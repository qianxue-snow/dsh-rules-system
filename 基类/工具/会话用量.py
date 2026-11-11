"""会话用量 —— 量当前这个对话烧了多少上下文。

为什么需要它
------------
模型自己**看不到**上下文用了多少（DSH 没有暴露这类环境变量）。
但每次模型调用的 `usage` 都写在自己的会话日志里，所以可以**事后读出来**。

它解决的事
----------
上下文是**累加**的：每轮都要重新处理全部历史，成本随轮次近似二次增长。
实测一个长会话能到 50~68 万 token；而开新会话只要 2~3 万。
**所以"该换会话了"这件事，越早越好，但要有个依据。**

用法
----
    python 会话用量.py                    # 打印用量
    python 会话用量.py --threshold 400000 # 超阈值退出码 1（给脚本用）
    python 会话用量.py --quiet            # 只打一行
    python 会话用量.py --selftest         # 自检（先校准尺子）

退出码
------
    0 = 正常（未超阈值 / 没给阈值）
    1 = 超过阈值
    2 = 环境不满足（找不到会话日志等）

依赖
----
    python -m pip install zstandard

口径
----
    context = inputTokens + cacheReadTokens
    （cache 命中部分同样占上下文窗口，所以必须一起算）
"""
from __future__ import annotations

import argparse
import glob
import json
import os
import sys


def _fix_console() -> None:
    """Windows 控制台默认 GBK，中文会变成乱码（实测把整个报告打成 `�`）。

    把输出编码和代码页都切成 UTF-8。全部 try 掉：失败只影响显示，不影响数字。
    """
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    if os.name == "nt":
        try:
            import ctypes

            ctypes.windll.kernel32.SetConsoleOutputCP(65001)
        except Exception:
            pass


def find_log(session_id: str) -> str | None:
    """在 ~/.dsh/sessions/*/session-<id>/ 下找日志。"""
    home = os.environ.get("DSH_HOME") or os.path.expanduser("~/.dsh")
    base = os.path.join(home, "sessions")
    if not os.path.isdir(base):
        return None
    for p in glob.glob(os.path.join(base, "*", session_id, "*.zstd")):
        return p
    return None


def read_usage(log_path: str) -> list[dict]:
    """解压日志，取出所有带 usage 的 assistant/message。"""
    import zstandard  # 延迟导入：--selftest 不依赖它

    with open(log_path, "rb") as fh:
        raw = zstandard.ZstdDecompressor().stream_reader(fh).read()
    text = raw.decode("utf-8", "replace")

    out: list[dict] = []
    for line in text.split("\n"):
        line = line.strip()
        if not line or "usage" not in line:
            continue
        try:
            ev = json.loads(line)
        except Exception:
            continue
        if ev.get("type") != "assistant/message":
            continue
        u = (ev.get("data") or {}).get("usage") or {}
        if u:
            out.append(u)
    return out


def ctx_of(u: dict) -> int:
    """上下文大小 = 非缓存输入 + 缓存命中。"""
    return int(u.get("inputTokens") or 0) + int(u.get("cacheReadTokens") or 0)


def selftest() -> int:
    """自检：造已知样本，确认统计口径对。"""
    _fix_console()
    checks = []

    def ck(name, got, want):
        ok = got == want
        checks.append(ok)
        print(f"  {'OK  ' if ok else 'FAIL'} {name}: {got}" + ("" if ok else f"（应为 {want}）"))

    # 样本 1：context = input + cache
    ck("context 计算", ctx_of({"inputTokens": 184, "cacheReadTokens": 524800}), 524984)
    # 样本 2：缺字段不炸
    ck("缺字段", ctx_of({}), 0)
    ck("只有 cache", ctx_of({"cacheReadTokens": 100}), 100)
    # 样本 3：峰值取自最大值
    seq = [{"inputTokens": 10, "cacheReadTokens": 100},
           {"inputTokens": 10, "cacheReadTokens": 900},
           {"inputTokens": 5, "cacheReadTokens": 50}]
    ck("峰值", max(ctx_of(x) for x in seq), 910)
    # 样本 4：最新取自最后一个
    ck("最新", ctx_of(seq[-1]), 55)

    n = sum(1 for x in checks if x)
    print(f"\n自检 {'通过' if n == len(checks) else '失败'}（{n}/{len(checks)} 项）")
    return 0 if n == len(checks) else 1


def main() -> int:
    _fix_console()
    ap = argparse.ArgumentParser(add_help=True)
    ap.add_argument("--threshold", type=int, default=0, help="超过这个上下文大小就返回退出码 1")
    ap.add_argument("--quiet", action="store_true", help="只打一行")
    ap.add_argument("--selftest", action="store_true", help="跑自检，不读日志")
    args = ap.parse_args()

    if args.selftest:
        return selftest()

    sid = os.environ.get("DSH_SESSION_ID", "").strip()
    if not sid:
        print("[会话用量] 环境里没有 DSH_SESSION_ID —— 不在 DSH 会话里跑？", file=sys.stderr)
        return 2

    log = find_log(sid)
    if not log:
        print(f"[会话用量] 找不到会话日志：{sid}", file=sys.stderr)
        print(f"           一般在 {os.environ.get('DSH_HOME', '~/.dsh')}\\sessions\\*\\{sid}\\", file=sys.stderr)
        return 2

    try:
        us = read_usage(log)
    except ImportError:
        print("[会话用量] 缺依赖：python -m pip install zstandard", file=sys.stderr)
        return 2
    except Exception as exc:
        print(f"[会话用量] 读日志失败：{type(exc).__name__}: {exc}", file=sys.stderr)
        return 2

    if not us:
        print("[会话用量] 日志里没有 usage 记录（会话刚开始？）", file=sys.stderr)
        return 2

    last = ctx_of(us[-1])
    peak = max(ctx_of(u) for u in us)
    size_mb = os.path.getsize(log) / 1024 / 1024

    if args.quiet:
        line = f"上下文 {last / 10000:.1f} 万（峰值 {peak / 10000:.1f} 万，{len(us)} 次调用）"
        if args.threshold:
            line += "  ⚠ 超阈值，建议换会话" if last > args.threshold else "  ✓ 未超阈值"
        print(line)
    else:
        print(f"会话 {sid}")
        print(f"  日志      {log}")
        print(f"             {size_mb:.1f} MB（压缩）")
        print(f"  调用次数  {len(us)}")
        print(f"  当前上下文 {last:>9,} token  （{last / 10000:.1f} 万）")
        print(f"  全程峰值   {peak:>9,} token  （{peak / 10000:.1f} 万）")
        if args.threshold:
            if last > args.threshold:
                print(f"  阈值       {args.threshold:>9,}  →  ⚠ 超了，建议换会话")
            else:
                print(f"  阈值       {args.threshold:>9,}  →  ✓ 未超")
        else:
            print("  提示      上下文是累加的，越拖越贵。交接只要几百 token，摸索要几万。")

    if args.threshold and last > args.threshold:
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
