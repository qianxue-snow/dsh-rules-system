<#
.SYNOPSIS
  check-md.ps1 —— 检查 Markdown 的结构问题（重复章节、代码块未闭合、表格列数、没填的占位符）。

.DESCRIPTION
  查的是**结构**，不查内容对不对：
    - 章节编号重复（同级别同编号出现两次，如实测过的两个「## 四、」）
    - 章节标题重复（同一节标题文字出现两次，可能是替换没删干净）
    - 代码块 ``` 是否成对
    - 表格每行列数是否一致（含内联代码的行跳过，避免把 \| 当分隔符）
    - 模板文件里没填的占位符（〔…〕）
    - 空章节（标题下面没有实际内容）

  为什么需要它：改长文档最常见的失败是**替换锚点写窄了、旧内容留在原地**——
  文件合法、工具不报错，只是内容重复了。那类问题靠"记特征串找回来"查不出（那是查"少了"），
  必须靠**穷举结构**来查。

.PARAMETER Path
  文件、目录或通配符。默认检查规则系统目录下所有 .md。

.PARAMETER Template
  按**模板**口径检查：允许空章节（骨架本来就留空等人填），并检查占位符。
  （路径里含 `骨架\`、或文件是 `_本工作区.md` / `_同步.md` 时自动启用。）

.PARAMETER Quiet
  只输出有问题的文件。

.EXAMPLE
  .\check-md.ps1
  .\check-md.ps1 -Path .\规则系统\骨架 -Template
  .\check-md.ps1 -Path C:\path\to\file.md
#>
[CmdletBinding()]
param(
    [string]$Path,
    [switch]$Template,
    [switch]$Strict,
    [switch]$SelfTest,
    [switch]$IncludeArchive,
    [switch]$Quiet
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------- 跳过规则
# 历史快照**不该被改**：它们是当时的事实记录，里面"空章节""旧编号"是真实存档状态。
# 对它们报问题是纯噪音——**工具不该为不该管的东西制造噪音**。
# 判据：路径里出现这些标记就跳过（除非用 -IncludeArchive 明确要求查）。
$SkipPattern = '\\\.git\\|\\__pycache__\\|\\_backup[^\\]*\\|\\archive\\|\.bak\.md$|\.bak$'
if ($IncludeArchive) { $SkipPattern = '\\\.git\\|\\__pycache__\\' }

# ---------------------------------------------------------------- 收集文件
$files = @()
if ($Path) {
    if (Test-Path -LiteralPath $Path -PathType Container) {
        $files = Get-ChildItem -LiteralPath $Path -Recurse -File -Filter '*.md' |
                 Where-Object { $_.FullName -notmatch $SkipPattern }
    } elseif (Test-Path -LiteralPath $Path -PathType Leaf) {
        $files = @(Get-Item -LiteralPath $Path)     # 直接点名的文件，一律检查
    } else {
        $files = Get-ChildItem -Path $Path -File -ErrorAction SilentlyContinue
    }
} else {
    $root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    $files = Get-ChildItem -LiteralPath $root -Recurse -File -Filter '*.md' |
             Where-Object { $_.FullName -notmatch $SkipPattern }
}

if (-not $files) { Write-Host "没有找到 .md 文件。" -ForegroundColor Yellow; exit 1 }

# ---------------------------------------------------------------- 检查一个文件
function Test-MdFile([string]$file) {
    $lines = Get-Content -LiteralPath $file -Encoding UTF8
    $issues = [System.Collections.Generic.List[string]]::new()
    $inTemplate = $Template -or ($file -match '\\骨架\\') -or ($file -match '_本工作区\.md$') -or ($file -match '_同步\.md$')

    # ---- 第一遍：围栏、标题、表格
    $inFence = $false
    $fenceOpenLine = 0
    $headings = @()
    $tableRows = @()
    $tableStart = 0
    $tables = @()

    for ($i = 0; $i -lt $lines.Count; $i++) {
        $ln = $lines[$i]; $no = $i + 1

        if ($ln -match '^\s*```') {
            if ($inFence) { $inFence = $false } else { $inFence = $true; $fenceOpenLine = $no }
            if ($tableRows.Count) { $tables += [pscustomobject]@{ Start=$tableStart; Rows=@($tableRows) }; $tableRows=@() }
            continue
        }
        if ($inFence) { continue }

        if ($ln -match '^(#{1,6})\s+(.+?)\s*$') {
            $level = $Matches[1].Length
            $text  = $Matches[2]
            $num = $null
            if ($text -match '^(\d+(?:\.\d+)*)[、.．·]?\s+') { $num = $Matches[1] }
            elseif ($text -match '^([〇零一二三四五六七八九十百]+)、') { $num = $Matches[1] }
            elseif ($text -match '^(\d+)、\s*') { $num = $Matches[1] }
            $headings += [pscustomobject]@{ Line=$no; Level=$level; Text=$text; Num=$num }
        }

        if ($ln -match '^\s*\|.*\|\s*$') {
            if ($tableRows.Count -eq 0) { $tableStart = $no }
            $tableRows += $ln
        } else {
            if ($tableRows.Count) { $tables += [pscustomobject]@{ Start=$tableStart; Rows=@($tableRows) }; $tableRows=@() }
        }
    }
    if ($inFence) { $issues.Add("代码块未闭合：第 $fenceOpenLine 行开了 ```` 一直没关") }
    if ($tableRows.Count) { $tables += [pscustomobject]@{ Start=$tableStart; Rows=@($tableRows) } }

    # ---- 章节编号重复 + 递增检查
    # 递增检查抓的是"替换锚点写窄、旧内容留在原地"这类：旧段落的编号会插到中间。
    $seenNum = @{}
    $lastNum = @{}          # level -> 上一次的编号
    foreach ($h in $headings) {
        if (-not $h.Num) { continue }
        $key = "$($h.Level)|$($h.Num)"
        if ($seenNum.ContainsKey($key)) {
            $issues.Add("章节编号重复：第 $($seenNum[$key]) 行和第 $($h.Line) 行都是 级别$($h.Level)「$($h.Num)」 → `"$($h.Text)`"")
        } else { $seenNum[$key] = $h.Line }

        if ($lastNum.ContainsKey($h.Level)) {
            $prev = $lastNum[$h.Level]
            if ($prev.Text -match '^\d' -and $h.Num -match '^\d') {
                if ([double]$h.Num -le [double]$prev.Num) {
                    $issues.Add("章节编号没有递增：第 $($h.Line) 行「$($h.Num)」不大于第 $($prev.Line) 行「$($prev.Num)」 —— 疑似旧内容残留")
                }
            }
        }
        $lastNum[$h.Level] = [pscustomobject]@{ Num=$h.Num; Line=$h.Line; Text=$h.Text }
    }

    # ---- 同级同名子标题重复（替换锚点写窄、旧段落残留的典型特征）
    # 判据：同一个父章节下面，两个**同级**标题文字完全相同。
    # 例：一个 ## 附 下面出现两个 ### 步骤 —— 实测就是这么漏的。
    $parentStack = @{}
    $childSeen = @{}
    for ($k = 0; $k -lt $headings.Count; $k++) {
        $h = $headings[$k]
        # 找到当前标题的父章节（前面最近的、级别更小的标题）
        $parentKey = ''
        for ($p = $k - 1; $p -ge 0; $p--) {
            if ($headings[$p].Level -lt $h.Level) { $parentKey = "$($headings[$p].Line)"; break }
        }
        $t = ($h.Text -replace '[*_`]', '').Trim()
        $ck = "$parentKey|$($h.Level)|$t"
        if ($childSeen.ContainsKey($ck)) {
            $issues.Add("同级同名子标题重复：第 $($childSeen[$ck]) 行和第 $($h.Line) 行，同一个父章节下都叫「$t」——疑似旧内容残留")
        } else {
            $childSeen[$ck] = $h.Line
        }
    }

    # ---- 章节标题重复（跳过纯编号标题，只报 6 字以上）
    # 注意：跨版本的父标题下常常有同名子标题（如每个版本都有一个"这一版要治什么"），
    # 那是合法的。所以这一项默认只做提示，不算失败——用 -Strict 才当问题。
    $seenText = @{}
    foreach ($h in $headings) {
        if ($h.Num) { continue }
        $t = ($h.Text -replace '[*_`]', '').Trim()
        if ($t.Length -lt 6) { continue }
        if ($seenText.ContainsKey($t)) {
            if ($Strict) {
                $issues.Add("章节标题重复：第 $($seenText[$t]) 行和第 $($h.Line) 行都叫「$t」")
            }
        } else { $seenText[$t] = $h.Line }
    }

    # ---- 表格列数（跳过分隔行与含内联代码的行）
    foreach ($tb in $tables) {
        $counts = @(); $skip = $false
        foreach ($row in $tb.Rows) {
            if ($row -match '^\s*\|[\s\-:|]+\|\s*$') { continue }
            if ($row -match '`') { $skip = $true; continue }
            $counts += (($row.Trim().Trim('|') -split '\|')).Count
        }
        if (-not $skip -and $counts.Count -gt 1) {
            $uniq = $counts | Sort-Object -Unique
            if ($uniq.Count -gt 1) {
                $issues.Add("表格列数不一致：第 $($tb.Start) 行起，各行为 $($counts -join '/') 列")
            }
        }
    }

    # ---- 占位符（模板口径才查）
    if ($inTemplate) {
        $f2 = $false
        for ($i = 0; $i -lt $lines.Count; $i++) {
            $ln = $lines[$i]
            if ($ln -match '^\s*```') { $f2 = -not $f2; continue }
            if ($f2) { continue }
            if ($ln -match '^\s*>') { continue }
            $stripped = $ln -replace '`[^`]*`', ''
            foreach ($m in [regex]::Matches($stripped, '〔[^〕]*〕')) {
                $v = $m.Value
                if ($v -eq '〔〕' -or $v -eq '〔…〕') { continue }
                if ($v -match '〔例[:：]') { continue }
                $issues.Add("占位符没填：第 $($i+1) 行 → $v")
            }
        }
    }

    # ---- 空章节（模板口径不查）
    # 只查**叶子章节**（后面紧跟的是同级或更高级标题）。
    # 有子标题的父章节不算空——内容在子标题下面，这是正常写法。
    if (-not $inTemplate) {
        for ($k = 0; $k -lt $headings.Count; $k++) {
            $cur = $headings[$k]
            $isLeaf = $true
            for ($n = $k + 1; $n -lt $headings.Count; $n++) {
                if ($headings[$n].Level -le $cur.Level) { break }
                $isLeaf = $false; break     # 下一节更深 → 这是父章节
            }
            if (-not $isLeaf) { continue }

            $stop = if ($k + 1 -lt $headings.Count) { $headings[$k+1].Line - 1 } else { $lines.Count }
            $hasBody = $false
            for ($j = $cur.Line; $j -lt $stop; $j++) {
                $c = $lines[$j].Trim()
                if (-not $c) { continue }
                if ($c -eq '---') { continue }
                $hasBody = $true; break
            }
            if (-not $hasBody) {
                $issues.Add("空章节：第 $($cur.Line) 行「$($cur.Text)」下面没有内容")
            }
        }
    }

    return [pscustomobject]@{ File=$file; Lines=$lines.Count; Issues=$issues }
}

# ---------------------------------------------------------------- 自检
# 原则：**先证明尺子是准的，再用它量。**
# 造已知答案的样本 → 该报的必须报，该过必须过 → 两条都满足才算这个检查器可信。
function Invoke-SelfTest {
    Write-Host "=== 自检：用已知答案的样本验证检查器本身 ===" -ForegroundColor Cyan
    $tmp = Join-Path $env:TEMP ("checkmd-selftest-" + [guid]::NewGuid().ToString('N') + ".md")
    $pass = 0; $fail = 0

    # 样本 1：含 3 种已知 bug，必须全部报出来
    $bad = @(
        '# 样本', '', '## 一、第一节', '', '内容一。', '',
        '## 二、第二节', '', '内容二。', '',
        '## 二、编号重复', '', '内容三。', '',           # bug 1：编号重复
        '### 步骤', '', '旧内容。', '',
        '### 步骤', '', '又一份。', '',                   # bug 2：同级同名子标题
        '~~~'                                              # bug 3：代码块未闭合
    ) -join "`n"
    $bad = $bad -replace '~~~', '```'
    [System.IO.File]::WriteAllText($tmp, $bad, (New-Object System.Text.UTF8Encoding($false)))

    $r = Test-MdFile $tmp
    $found = ($r.Issues | Out-String)
    foreach ($nm in @('章节编号重复', '同级同名子标题重复', '代码块未闭合')) {
        if ($found -match [regex]::Escape($nm)) { $pass++; Write-Host "  OK   抓到「$nm」" -ForegroundColor Green }
        else { $fail++; Write-Host "  FAIL 没抓到「$nm」" -ForegroundColor Red }
    }

    # 样本 2：干净文件，必须零误报
    $good = @(
        '# 干净样本', '', '## 一、第一节', '', '内容一。', '',
        '## 二、第二节', '', '内容二。', '',
        '### 子节', '', '子内容。', '',
        '~~~', 'code', '~~~'
    ) -join "`n"
    $good = $good -replace '~~~', '```'
    [System.IO.File]::WriteAllText($tmp, $good, (New-Object System.Text.UTF8Encoding($false)))
    $r2 = Test-MdFile $tmp
    if ($r2.Issues.Count -eq 0) { $pass++; Write-Host "  OK   干净文件零误报" -ForegroundColor Green }
    else { $fail++; Write-Host ("  FAIL 干净文件误报 {0} 处：{1}" -f $r2.Issues.Count, ($r2.Issues -join '; ')) -ForegroundColor Red }

    # 样本 3：有子标题的父章节不算空
    $nested = @(
        '# 样本', '', '## 一、父章节', '', '### 1. 子节', '', '内容。', ''
    ) -join "`n"
    [System.IO.File]::WriteAllText($tmp, $nested, (New-Object System.Text.UTF8Encoding($false)))
    $r3 = Test-MdFile $tmp
    $emptyIssue = @($r3.Issues | Where-Object { $_ -match '空章节' })
    if ($emptyIssue.Count -eq 0) { $pass++; Write-Host "  OK   有子标题的父章节没被误报为空" -ForegroundColor Green }
    else { $fail++; Write-Host "  FAIL 父章节被误报为空" -ForegroundColor Red }

    Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue

    Write-Host ""
    if ($fail -eq 0) {
        Write-Host ("自检通过（{0} 项）。这个检查器可以信。" -f $pass) -ForegroundColor Green
        return 0
    } else {
        Write-Host ("自检失败（{0} 通过 / {1} 失败）。**检查器本身不可信，它报的结果不要拿来下结论。**" -f $pass, $fail) -ForegroundColor Red
        return 1
    }
}

if ($SelfTest) { exit (Invoke-SelfTest) }

# ---------------------------------------------------------------- 跑
$totalIssues = 0; $badFiles = 0
foreach ($f in $files) {
    $r = Test-MdFile $f.FullName
    if ($r.Issues.Count -eq 0) {
        if (-not $Quiet) { Write-Host ("OK  {0}  ({1} 行)" -f $f.Name, $r.Lines) -ForegroundColor Green }
    } else {
        $badFiles++; $totalIssues += $r.Issues.Count
        Write-Host ("XX  {0}  ({1} 行)  {2} 个问题" -f $r.File, $r.Lines, $r.Issues.Count) -ForegroundColor Yellow
        foreach ($iss in $r.Issues) { Write-Host ("    - {0}" -f $iss) -ForegroundColor Gray }
    }
}

Write-Host ""
if ($totalIssues -eq 0) {
    Write-Host ("全部通过：{0} 个文件" -f $files.Count) -ForegroundColor Green
    exit 0
} else {
    Write-Host ("{0} 个文件有问题，共 {1} 处" -f $badFiles, $totalIssues) -ForegroundColor Yellow
    Write-Host "提示：先把这些当线索，别当结论。**先跑 -SelfTest 确认检查器自己没坏。**" -ForegroundColor DarkGray
    exit 1
}
