# -*- coding: utf-8 -*-
"""verify_docx.py —— 自检生成的 .docx：不打开 Word，直接读结构。

检查项：
  1. 包结构 —— 必需部件是否齐全、zip 是否损坏
  2. XML 合法性 —— 每个部件都要能被解析
  3. 内容完整性 —— 与同名 .md 比对中文字数与表格数
  4. OOXML 规范性 —— 子元素顺序、禁止相邻表格（Word 对顺序是强校验）
  5. 关键串断言 —— 用 --expect 传入必须出现的词

为什么不用 Word 打开来验证：启动 Word 要 20+ 秒，而且"能打开"证明不了内容对。
本脚本毫秒级完成，且能查出 Word 不会报的 OOXML 顺序错误。
"""

import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
REQUIRED = ['[Content_Types].xml', '_rels/.rels', 'word/document.xml',
            'word/_rels/document.xml.rels', 'word/styles.xml',
            'docProps/core.xml', 'docProps/app.xml']

# 必须出现的关键串；由 --expect 填充，默认为空（即跳过该项）
EXPECT = []


def cjk(s):
    return len(re.findall(r'[\u4e00-\u9fff]', s))


def check(out_dir):
    md_files = sorted(f for f in os.listdir(out_dir) if f.endswith('.md'))
    bad = 0
    for md_name in md_files:
        stem = os.path.splitext(md_name)[0]
        docx = os.path.join(out_dir, stem + '.docx')
        md = os.path.join(out_dir, md_name)
        label = stem[:22]
        problems = []

        if not os.path.exists(docx):
            print('FAIL %-24s 缺 .docx' % label)
            bad += 1
            continue

        with zipfile.ZipFile(docx) as z:
            names = z.namelist()
            missing = [p for p in REQUIRED if p not in names]
            if missing:
                problems.append('缺部件 %s' % missing)
            if z.testzip() is not None:
                problems.append('zip 损坏')
            # 每个 XML 部件都要能解析
            for p in names:
                if p.endswith('.xml') or p.endswith('.rels'):
                    try:
                        ET.fromstring(z.read(p))
                    except Exception as e:
                        problems.append('%s XML 非法: %s' % (p, e))
            doc = z.read('word/document.xml').decode('utf-8')

        root = ET.fromstring(doc)
        texts = [t.text or '' for t in root.iter(W + 't')]
        body_text = ''.join(texts)

        src = open(md, encoding='utf-8').read()
        src_cjk = cjk(src)
        out_cjk = cjk(body_text)
        # 允许微小差异：分隔行/标题标记符号
        if abs(src_cjk - out_cjk) > max(3, int(src_cjk * 0.01)):
            problems.append('字数不符 md=%d docx=%d' % (src_cjk, out_cjk))

        n_par = len(list(root.iter(W + 'p')))
        n_tbl = len(list(root.iter(W + 'tbl')))
        md_tbl = len(re.findall(r'^\|[\s\-:|]+\|$', src, re.M))
        if n_tbl != md_tbl:
            problems.append('表格数不符 md=%d docx=%d' % (md_tbl, n_tbl))

        # 相邻表格之间必须有段落（OOXML 结构要求）
        if '</w:tbl><w:tbl>' in doc or '<w:tbl><w:tbl>' in doc:
            problems.append('出现相邻表格')

        # 元素顺序断言（Word 对 OOXML 子元素顺序是强校验）
        RANK_P = {'pBdr': 1, 'shd': 2, 'spacing': 3, 'ind': 4, 'jc': 5, 'rPr': 6,
                  'rFonts': 7, 'b': 8, 'bCs': 9, 'color': 10, 'sz': 11, 'szCs': 12}
        RANK_TBLPR = {'tblW': 1, 'jc': 2, 'tblBorders': 3, 'tblLayout': 4, 'tblCellMar': 5}
        RANK_TCPR = {'tcW': 1, 'shd': 2, 'vAlign': 3}
        RANK_TRPR = {'trHeight': 1, 'tblHeader': 2}
        TAG_RE = re.compile(r'<w:([A-Za-z]+)[ />]')

        def order_ok(block, rank):
            seq = [rank[t] for t in TAG_RE.findall(block) if t in rank]
            return seq == sorted(seq), seq

        checks = [(r'<w:pPr>(.*?)</w:pPr>', RANK_P, 'pPr'),
                  (r'<w:tblPr>(.*?)</w:tblPr>', RANK_TBLPR, 'tblPr'),
                  (r'<w:tcPr>(.*?)</w:tcPr>', RANK_TCPR, 'tcPr'),
                  (r'<w:trPr>(.*?)</w:trPr>', RANK_TRPR, 'trPr')]
        for pat, rank, nm in checks:
            for blk in re.findall(pat, doc, re.S):
                ok, seq = order_ok(blk, rank)
                if not ok:
                    problems.append('%s 子元素顺序错：%s' % (nm, seq))
                    break

        # 关键串是否还在（用 --expect 传入；不传则跳过这项）
        for key in EXPECT:
            if key not in body_text:
                problems.append('丢了关键内容：%s' % key)

        # 元素顺序抽查：每个 rPr 里 rFonts 必须在 b/color/sz 之前
        for rpr in re.findall(r'<w:rPr>.*?</w:rPr>', doc):
            order = re.findall(r'<w:(rFonts|b|bCs|color|sz|szCs)\b', rpr)
            ideal = [x for x in ['rFonts', 'b', 'bCs', 'color', 'sz', 'szCs'] if x in order]
            if order != ideal:
                problems.append('rPr 顺序错：%s' % order)
                break

        status = 'PASS' if not problems else 'FAIL'
        if problems:
            bad += 1
        print('%s %-24s 段落=%-4d 表格=%d 字数 md=%d/docx=%d  %s'
              % (status, label, n_par, n_tbl, src_cjk, out_cjk,
                 '; '.join(problems) if problems else ''))
    print('\n结果：%d 个文件，%d 个有问题' % (len(md_files), bad))
    return 1 if bad else 0


if __name__ == '__main__':
    argv = sys.argv[1:]
    if argv and argv[0] in ('-h', '--help'):
        print(__doc__.strip())
        print()
        print('用法: python verify_docx.py [目录] [--expect 关键词1 关键词2 ...]')
        print('  对目录下每个 .md，检查同名 .docx 是否结构合法、内容完整。')
        print('  --expect 用来断言必须出现的关键串（如活动名、地点）。')
        print('  不指定目录则用当前目录。退出码 0=全部通过，1=有问题。')
        sys.exit(0)
    out_dir = '.'
    if argv and not argv[0].startswith('-'):
        out_dir = argv[0]
        argv = argv[1:]
    if '--expect' in argv:
        EXPECT = [k for k in argv[argv.index('--expect') + 1:] if not k.startswith('-')]
    sys.exit(check(out_dir))
