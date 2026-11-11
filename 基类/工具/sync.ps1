<#
.SYNOPSIS
  sync.ps1 —— 检查已应用模版的工作区，哪些落后于规则系统母版，并告诉你该做什么。

.DESCRIPTION
  它做四件事：
    1. 读出工作区 _同步.md 里记的版本号
    2. 按版本号从模版 更新日志.md 里挑出「工作区还没跟」的条目
    3. 把骨架里的「复制型」文件和母版逐一对比，看结构差了多少
    4. 输出一张行动清单

  只读，不改任何文件。要怎么改由你决定。

.PARAMETER Workspace
  工作区根目录。不传则读 已应用工作区.md 里的清单。

.PARAMETER All
  检查清单里的全部工作区。

.EXAMPLE
  .\sync.ps1                                  # 检查清单里所有工作区
  .\sync.ps1 -Workspace C:\path\to\ws         # 只检查一个
#>
[CmdletBinding()]
param(
    [string]$Workspace,
    [switch]$All
)

$ErrorActionPreference = 'Stop'
$RulesRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent   # 基类\工具 -> 规则系统
$Changelog = Join-Path $RulesRoot '更新日志.md'
$Skeleton  = Join-Path $RulesRoot '骨架'
$Registry  = Join-Path $RulesRoot '已应用工作区.md'

if (-not (Test-Path $Changelog)) { throw "找不到更新日志：$Changelog" }
if (-not (Test-Path $Skeleton))  { throw "找不到骨架目录：$Skeleton" }

# ---------------------------------------------------------------- 先拉新版本
# 要紧：本地 tag 只是"上次 fetch/clone 那一刻"的快照。
# 不先 fetch，就永远看不到母版发了新版——那"检查更新"这件事根本不会发生。
#
# 两个坑要绕：
#   1. 全局可能有 url.insteadOf 把 github.com 重写成**只读镜像**（如 gitclone.com）。
#      实测：`-c url.<镜像>.insteadof=` 传空**顶不掉**它（git 仍走镜像）。
#      所以改成"自己从 remote 算出真实地址"，显式把那个地址交给 fetch。
#   2. 国内直连 github 常常不通。若检测到常见本地代理端口在监听，就走代理。
$FetchNote = ''
$ProxyPorts = @(7897, 7890, 7891, 10809, 10808, 1080)   # Clash Verge 默认 7897
$Proxy = $null
foreach ($p in $ProxyPorts) {
    if (Get-NetTCPConnection -LocalPort $p -State Listen -ErrorAction SilentlyContinue) {
        $Proxy = "http://127.0.0.1:$p"
        break
    }
}

# 把被镜像重写过的地址还原成真实 GitHub 地址
function Get-CanonicalUrl([string]$u) {
    if ([string]::IsNullOrWhiteSpace($u)) { return $u }
    $t = $u.Trim()
    if ($t -match '^https?://[^/]*gitclone\.com/github\.com/(.+)$')     { return "https://github.com/$($Matches[1])" }
    if ($t -match '^https?://[^/]*gh-proxy\.com/https://github\.com/(.+)$') { return "https://github.com/$($Matches[1])" }
    return $t
}

$hasRemote = & git -C $RulesRoot remote 2>$null
if ($hasRemote) {
    $rawUrl   = (& git -C $RulesRoot remote get-url origin 2>$null | Select-Object -First 1)
    $realUrl  = Get-CanonicalUrl $rawUrl
    $mirrored = ($realUrl -ne $rawUrl)

    Write-Host "先从远程拉取 tag ……" -ForegroundColor DarkGray
    # 镜像重写只在"拉取失败、需要解释原因"时才提；成功时提它是纯噪音。

    $gitArgs = @('-C', $RulesRoot)
    if ($Proxy) {
        Write-Host "  （走代理 $Proxy）" -ForegroundColor DarkGray
        $gitArgs += @('-c', "http.proxy=$Proxy", '-c', "https.proxy=$Proxy")
    }
    $gitArgs += @('fetch', $realUrl, '--tags', '--prune')

    $oldEAP = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'      # git 的进度/stderr 不该让脚本崩
    $fetchOut = & git @gitArgs 2>&1
    $fetchCode = $LASTEXITCODE
    $ErrorActionPreference = $oldEAP

    if ($fetchCode -ne 0) {
        # 拉取失败**不是故障**：单机模式下母版就在本地，本地 tag 就是权威。
        # 所以这里只给一行灰字，不报红、不占篇幅。
        $FetchNote = "远程拉取没成功（多半是没挂代理）。本地 tag 就是权威，照常比对。"
        Write-Host "  · 远程拉取没成功，用本地 tag 比对（单机模式下这完全正常）" -ForegroundColor DarkGray
        if ($mirrored) { Write-Host "    （远程地址被镜像规则改写过，已改用：$realUrl）" -ForegroundColor DarkGray }
        if (-not $Proxy) { Write-Host "    （挂了代理会自动重试；不挂也不影响结果）" -ForegroundColor DarkGray }
    } else {
        Write-Host "  ✓ 已同步远程 tag" -ForegroundColor DarkGray
    }
} else {
    # 没有 remote 是**正常工作状态**，不是缺失。不要写成警告。
    Write-Host "  · 单机模式：母版就在本地，本地 tag 即最新，无需拉取" -ForegroundColor DarkGray
    $FetchNote = ''
}

# ---------------------------------------------------------------- 读版本
function Get-TagOrder {
    $tags = & git -C $RulesRoot tag 2>$null
    if (-not $tags) { return @() }
    $tags | Sort-Object { [version]($_ -replace '^v','') }
}

function Get-TagIndex([string]$tag, $order) {
    if ([string]::IsNullOrWhiteSpace($tag)) { return -1 }
    $t = $tag.Trim().Trim('`')
    $i = [array]::IndexOf($order, $t)
    return $i
}

$TagOrder = Get-TagOrder
if ($TagOrder.Count -eq 0) { throw "规则系统里没有 git tag，无法比较版本。先打 tag。" }
$LatestTag = $TagOrder[-1]
$LatestIdx = $TagOrder.Count - 1

# ---------------------------------------------------------------- 解析更新日志
# 只解析 <!-- ENTRIES --> 与 <!-- /ENTRIES --> 之间的内容，
# 这样文末的格式示例（那里也有一张同名表）不会被当成真条目。
# 解析时跳过 ``` 代码块，免得示例里写的标记被当真。
# 版本标题支持范围写法 `## 版本 v1.1 ~ v1.3`，取范围里最高的那个版本。
$entries = @()
$currentVer = ''
$inBlock = $false
$inFence = $false
foreach ($line in (Get-Content -LiteralPath $Changelog -Encoding UTF8)) {
    if ($line -match '^\s*```') { $inFence = -not $inFence; continue }
    if ($inFence) { continue }

    if ($line -match '<!--\s*ENTRIES\s*-->')  { $inBlock = $true;  continue }
    if ($line -match '<!--\s*/ENTRIES\s*-->') { $inBlock = $false; continue }
    if (-not $inBlock) { continue }

    if ($line -match '^##\s*版本\s*(v[\d.]+(?:\s*~\s*v[\d.]+)?)') {
        $vers = [regex]::Matches($Matches[1], 'v[\d.]+') | ForEach-Object { $_.Value }
        $currentVer = $vers | Sort-Object { [version]($_ -replace '^v','') } | Select-Object -Last 1
        continue
    }
    if ($line -match '^\|\s*(\d+)\s*\|\s*(.+?)\s*\|\s*(.+?)\s*\|\s*(.+?)\s*\|\s*$') {
        $entries += [pscustomobject]@{
            版本 = $currentVer
            序号 = [int]$Matches[1]
            变更 = $Matches[2]
            影响 = $Matches[3]
            建议 = $Matches[4]
        }
    }
}

# ---------------------------------------------------------------- 读工作区同步状态
function Read-SyncFile([string]$ws) {
    $f = Join-Path $ws '_同步.md'
    if (-not (Test-Path $f)) { return $null }
    $txt = Get-Content -LiteralPath $f -Encoding UTF8 -Raw
    $ver = ''
    if ($txt -match '模版版本\s*\|\s*`?(v[\d.]+)`?') { $ver = $Matches[1] }
    return [pscustomobject]@{ 路径 = $f; 版本 = $ver }
}

# ---------------------------------------------------------------- 复制型文件对比
# 只比"结构"：章节标题 + 表格行数。工作区填的内容本来就不一样，逐字比没意义。
function Get-Shape([string]$path) {
    if (-not (Test-Path $path)) { return $null }
    $lines = Get-Content -LiteralPath $path -Encoding UTF8
    $heads = $lines | Where-Object { $_ -match '^#{1,3}\s' } | ForEach-Object { $_.Trim() }
    $rows  = ($lines | Where-Object { $_ -match '^\|' }).Count
    return [pscustomobject]@{ 标题 = $heads; 表格行 = $rows; 行数 = $lines.Count }
}

# 骨架里所有需要"复制到工作区"的文件（README/_同步 除外）
$CopyFiles = Get-ChildItem -LiteralPath $Skeleton -File |
    Where-Object { $_.Name -ne 'README.md' } | Select-Object -ExpandProperty Name

# ---------------------------------------------------------------- 主流程
# $Type: 'copy'（从骨架复制的）| 'derived'（自己重写的，只借方法论与工具）
function Test-One([string]$ws, [string]$Type = 'copy') {
    $name = Split-Path $ws -Leaf
    $isDerived = ($Type -eq 'derived')
    Write-Host ""
    Write-Host ("=" * 68) -ForegroundColor DarkGray
    if ($isDerived) {
        Write-Host "  工作区：$name   （派生型）" -ForegroundColor Cyan
    } else {
        Write-Host "  工作区：$name   （复制型）" -ForegroundColor Cyan
    }
    Write-Host "  路径：  $ws" -ForegroundColor DarkGray
    Write-Host ("=" * 68) -ForegroundColor DarkGray

    if (-not (Test-Path $ws)) { Write-Host "  ✗ 目录不存在，跳过" -ForegroundColor Red; return }

    $sync = Read-SyncFile $ws
    if (-not $sync) {
        Write-Host "  ⚠ 没有 _同步.md —— 这个工作区没登记过同步状态。" -ForegroundColor Yellow
        Write-Host "    处理：从 骨架\_同步.md 复制一份到工作区根，填上当前版本号。" -ForegroundColor Gray
        $wsIdx = -1
    } else {
        $wsIdx = Get-TagIndex $sync.版本 $TagOrder
        if ($wsIdx -lt 0) {
            Write-Host "  ⚠ _同步.md 里的版本号 '$($sync.版本)' 在 git tag 里找不到。" -ForegroundColor Yellow
        } elseif ($wsIdx -eq $LatestIdx) {
            Write-Host "  ✓ 已是最新版本 $LatestTag" -ForegroundColor Green
        } else {
            Write-Host "  版本：$($sync.版本)  →  最新 $LatestTag" -ForegroundColor Yellow
        }
    }

    # --- 1. 待跟的更新条目
    $pending = @()
    if ($wsIdx -ge 0) {
        foreach ($e in $entries) {
            $eIdx = Get-TagIndex $e.版本 $TagOrder
            if ($eIdx -gt $wsIdx) { $pending += $e }
        }
    } else {
        $pending = $entries   # 没同步记录的，全都算待确认
    }

    $manual = @($pending | Where-Object { $_.建议 -match '要手动' })
    $auto   = @($pending | Where-Object { $_.建议 -match '自动生效' })
    $skip   = @($pending | Where-Object { $_.建议 -match '不用管' })

    Write-Host ""
    Write-Host "  待跟的更新条目：$($pending.Count) 条" -ForegroundColor White
    if ($auto.Count)   { Write-Host "    ✓ 自动生效 $($auto.Count) 条（不用操作）" -ForegroundColor Green }
    if ($skip.Count)   { Write-Host "    · 不用管   $($skip.Count) 条" -ForegroundColor DarkGray }
    if ($manual.Count) {
        Write-Host "    ⚠ 要手动   $($manual.Count) 条：" -ForegroundColor Yellow
        foreach ($m in $manual) {
            Write-Host ("        [{0} #{1}] {2}" -f $m.版本, $m.序号, $m.变更) -ForegroundColor Yellow
            Write-Host ("            影响：{0}" -f $m.影响) -ForegroundColor DarkGray
            Write-Host ("            做法：{0}" -f $m.建议) -ForegroundColor DarkGray
        }
    }

    # --- 2. 复制型文件的形状对比
    Write-Host ""
    if ($isDerived) {
        Write-Host "  骨架文件对比：**本工作区是派生型，这一节不适用，跳过**" -ForegroundColor DarkGray
        Write-Host "    （派生型自己重写了规则文件，和骨架没有对应关系。" -ForegroundColor DarkGray
        Write-Host "      它跟不跟母版，只看上面「待跟的更新条目」，不看结构对比。）" -ForegroundColor DarkGray
    } else {
        Write-Host "  复制型文件 vs 母版（只看结构，不看内容）：" -ForegroundColor White
        foreach ($cf in $CopyFiles) {
            $master = Join-Path $Skeleton $cf
            $mShape = Get-Shape $master
            # 找副本：AGENTS/本工作区/同步 在根，其余可能也在根或任务目录
            $cands = @(
                (Join-Path $ws $cf)
            ) + (Get-ChildItem -LiteralPath $ws -Recurse -Filter $cf -File -ErrorAction SilentlyContinue |
                    Select-Object -ExpandProperty FullName)
            $copy = $cands | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1

            if (-not $copy) {
                Write-Host ("    ?  {0,-22} 工作区里没找到" -f $cf) -ForegroundColor DarkGray
                continue
            }
            $cShape = Get-Shape $copy
            $missHead = @($mShape.标题 | Where-Object { $_ -notin $cShape.标题 })
            $extraHead = @($cShape.标题 | Where-Object { $_ -notin $mShape.标题 })

            if ($missHead.Count -eq 0 -and $extraHead.Count -eq 0 -and $mShape.表格行 -eq $cShape.表格行) {
                Write-Host ("    ✓  {0,-22} 结构一致" -f $cf) -ForegroundColor Green
            } else {
                Write-Host ("    ⚠  {0,-22} 结构有差异" -f $cf) -ForegroundColor Yellow
                foreach ($h in $missHead) { Write-Host ("         母版有、这边没有：{0}" -f $h) -ForegroundColor Yellow }
                foreach ($h in $extraHead) { Write-Host ("         这边有、母版没有：{0}" -f $h) -ForegroundColor DarkGray }
                if ($mShape.表格行 -ne $cShape.表格行) {
                    Write-Host ("         表格行数：母版 {0} / 这边 {1}" -f $mShape.表格行, $cShape.表格行) -ForegroundColor DarkGray
                }
            }
        }
    }

    # --- 3. 结论
    Write-Host ""
    if ($pending.Count -eq 0 -and $wsIdx -eq $LatestIdx) {
        Write-Host "  → 无需操作" -ForegroundColor Green
    } else {
        Write-Host "  → 要做的：" -ForegroundColor Cyan
        Write-Host "     1. 读 更新日志.md 里上面列出的「要手动」条目，逐条改" -ForegroundColor Cyan
        Write-Host "     2. 改完把 _同步.md 的版本号更新成 $LatestTag，并在「同步历史」记一行" -ForegroundColor Cyan
        Write-Host "     3. 故意不跟的，写进 _同步.md 的「已知偏离」" -ForegroundColor Cyan
    }
}

# ---------------------------------------------------------------- 选工作区
# 登记表一行可以是：
#   - <绝对路径>
#   - <绝对路径> | derived      ← 派生型：自己重写了规则，骨架对比对它不适用
#
# 用 -Workspace 单查时，类型默认不知道，就去登记表里找一下同一个路径——
# 找不到才当复制型。否则单查派生型工作区会误报一堆结构差异。
function Get-RegistryType([string]$path) {
    if (-not (Test-Path $Registry)) { return 'copy' }
    $inList = $false
    foreach ($line in (Get-Content -LiteralPath $Registry -Encoding UTF8)) {
        if ($line -match '<!--\s*LIST\s*-->')  { $inList = $true;  continue }
        if ($line -match '<!--\s*/LIST\s*-->') { $inList = $false; continue }
        if (-not $inList) { continue }
        if ($line -notmatch '^\s*-\s+\S') { continue }
        $l = $line -replace '^\s*-\s+',''
        $t = 'copy'
        if ($l -match '\|\s*(derived|派生)\s*$') { $t = 'derived'; $l = $l -replace '\|\s*(derived|派生)\s*$','' }
        $l = $l.Trim().Trim('`')
        if ($l -eq $path) { return $t }
    }
    return 'copy'
}

$targets = @()
if ($Workspace) {
    $targets = @([pscustomobject]@{ 路径 = $Workspace; 类型 = (Get-RegistryType $Workspace) })
} elseif (Test-Path $Registry) {
    # 只解析 <!-- LIST --> 与 <!-- /LIST --> 之间的行。
    # 不这么做的话，登记表里任何以 "- " 开头的说明文字都会被当成路径。
    $inList = $false
    $targets = Get-Content -LiteralPath $Registry -Encoding UTF8 |
        ForEach-Object {
            if ($_ -match '<!--\s*LIST\s*-->')  { $inList = $true;  return }
            if ($_ -match '<!--\s*/LIST\s*-->') { $inList = $false; return }
            if (-not $inList) { return }
            if ($_ -notmatch '^\s*-\s+\S') { return }
            $line = $_ -replace '^\s*-\s+',''
            $type = 'copy'
            if ($line -match '\|\s*(derived|派生)\s*$') {
                $type = 'derived'
                $line = $line -replace '\|\s*(derived|派生)\s*$',''
            }
            [pscustomobject]@{ 路径 = $line.Trim().Trim('`'); 类型 = $type }
        } |
        Where-Object { $_ -and $_.路径 -and (Test-Path $_.路径) }
    if (-not $targets) { Write-Host "已应用工作区.md 的 <!-- LIST --> 区块里没有有效路径。" -ForegroundColor Yellow }
} else {
    Write-Host "没有传 -Workspace，也没找到 已应用工作区.md。" -ForegroundColor Yellow
    Write-Host "用法：.\sync.ps1 -Workspace <工作区根目录>" -ForegroundColor Gray
    exit 1
}

foreach ($t in $targets) { Test-One $t.路径 $t.类型 }

Write-Host ""
if ($FetchNote) { Write-Host "注意：$FetchNote" -ForegroundColor Yellow }
Write-Host "（本脚本只读，没有改动任何文件）" -ForegroundColor DarkGray
