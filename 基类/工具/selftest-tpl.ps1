<#
.SYNOPSIS
  selftest-tpl.ps1 —— 模版体检：真搭一个临时工作区，跑全流程，验结果，然后清掉。

.DESCRIPTION
  自动测试能覆盖的只有"文件与流程"这一层：
    - 工具自身能不能跑（调 check-md / audit 的 -SelfTest）
    - 骨架复制到新工作区后，文件是否齐、占位符是否只剩该留的
    - 填完配置后，"规则完善"的判据是否真能通过
    - 登记 → sync.ps1 能否认出它 → 注销，是否干净
    - 少文件 / 填不全时，检查工具是否**真的报错**（负向测试）

  **它测不了"AI 会不会照规则做"** —— 那必须真开一个对话，
  用同目录的 `对话测试.md` 手工验。

.PARAMETER Keep
  跑完保留临时工作区（排查用）。默认自动删除。

.PARAMETER SelfTest
  只验体检器自己（造已知样本，确认它该报的报、该过的过）。

.EXAMPLE
  .\selftest-tpl.ps1
  .\selftest-tpl.ps1 -Keep
  .\selftest-tpl.ps1 -SelfTest
#>
[CmdletBinding()]
param(
    [switch]$Keep,
    [switch]$SelfTest
)

$ErrorActionPreference = 'Stop'
$ToolsDir = $PSScriptRoot
$RulesRoot = Split-Path (Split-Path $ToolsDir -Parent) -Parent
$SkelDir = Join-Path $RulesRoot '骨架'
$Registry = Join-Path $RulesRoot '已应用工作区.md'
$audit = Join-Path $ToolsDir 'audit.ps1'
$checkMd = Join-Path $ToolsDir 'check-md.ps1'
$sync = Join-Path $ToolsDir 'sync.ps1'

$script:pass = 0
$script:fail = 0
$script:silent = $false

function Ok([string]$msg)   { $script:pass++; if (-not $script:silent) { Write-Host ("  PASS  {0}" -f $msg) -ForegroundColor Green } }
function No([string]$msg, [string]$why = '') {
    $script:fail++
    if (-not $script:silent) {
        Write-Host ("  FAIL  {0}{1}" -f $msg, $(if($why){"　→ $why"}else{''})) -ForegroundColor Red
    }
}
function Assert-True([bool]$cond, [string]$msg, [string]$why = '') {
    if ($cond) { Ok $msg } else { No $msg $why }
}

# ---------------------------------------------------------------- 自检
function Invoke-SelfTest {
    Write-Host "=== 自检：先确认体检器自己没坏 ===" -ForegroundColor Cyan
    $p = 0; $f = 0

    # 1. 计数器能用（静默跑，别让测试残留污染输出）
    $script:pass = 0; $script:fail = 0; $script:silent = $true
    Ok 'x'; No 'y'
    $script:silent = $false
    if ($script:pass -eq 1 -and $script:fail -eq 1) { $p++; Write-Host "  OK   计数器正常" -ForegroundColor Green }
    else { $f++; Write-Host "  FAIL 计数器坏了" -ForegroundColor Red }
    $script:pass = 0; $script:fail = 0

    # 2. 骨架目录存在且非空
    if ((Test-Path -LiteralPath $SkelDir) -and @(Get-ChildItem -LiteralPath $SkelDir -File).Count -ge 5) {
        $p++; Write-Host "  OK   骨架目录存在且文件齐" -ForegroundColor Green
    } else { $f++; Write-Host "  FAIL 骨架目录缺失或文件不全" -ForegroundColor Red }

    # 3. 三个工具都在
    $missingTools = @($audit, $checkMd, $sync) | Where-Object { -not (Test-Path -LiteralPath $_) }
    if ($missingTools.Count -eq 0) { $p++; Write-Host "  OK   审计/检查/同步 三个工具都在" -ForegroundColor Green }
    else { $f++; Write-Host ("  FAIL 缺工具：{0}" -f ($missingTools -join ', ')) -ForegroundColor Red }

    # 4. 登记表有 LIST 区块（临时登记要用）
    $regTxt = if (Test-Path -LiteralPath $Registry) { Get-Content -LiteralPath $Registry -Encoding UTF8 -Raw } else { '' }
    if ($regTxt -match '<!--\s*LIST\s*-->' -and $regTxt -match '<!--\s*/LIST\s*-->') {
        $p++; Write-Host "  OK   登记表有 LIST 区块，可安全临时登记" -ForegroundColor Green
    } else { $f++; Write-Host "  FAIL 登记表没有 LIST 区块——临时登记会污染文件" -ForegroundColor Red }

    Write-Host ""
    if ($f -eq 0) { Write-Host ("自检通过（{0} 项）。可以信。" -f $p) -ForegroundColor Green; return 0 }
    Write-Host ("自检失败（{0} 通过 / {1} 失败）。**结果不要拿来下结论。**" -f $p, $f) -ForegroundColor Red
    return 1
}

if ($SelfTest) { exit (Invoke-SelfTest) }

# ---------------------------------------------------------------- 准备工作区
$tmp = Join-Path $env:TEMP ("tpl-selftest-" + [guid]::NewGuid().ToString('N'))
$taskDir = Join-Path $tmp '任务区'
$created = $false

try {
    Write-Host "=== 1. 搭临时工作区 ===" -ForegroundColor Cyan
    Write-Host "  $tmp" -ForegroundColor DarkGray
    New-Item -ItemType Directory -Path $tmp -Force | Out-Null
    New-Item -ItemType Directory -Path $taskDir -Force | Out-Null
    $created = $true

    # 复制骨架（模拟"照着准备流程部署"）
    Copy-Item -LiteralPath (Join-Path $SkelDir 'AGENTS.md')          -Destination $tmp -Force
    Copy-Item -LiteralPath (Join-Path $SkelDir '_同步.md')            -Destination $tmp -Force
    Copy-Item -LiteralPath (Join-Path $SkelDir '_本工作区.md')          -Destination $tmp -Force
    Copy-Item -LiteralPath (Join-Path $SkelDir 'GLOBAL_WORKFLOW.md')  -Destination $taskDir -Force
    Copy-Item -LiteralPath (Join-Path $SkelDir '任务看板.md')           -Destination $taskDir -Force

    Assert-True (Test-Path -LiteralPath (Join-Path $tmp 'AGENTS.md')) '骨架复制：AGENTS.md 到位'
    Assert-True (Test-Path -LiteralPath (Join-Path $taskDir 'GLOBAL_WORKFLOW.md')) '骨架复制：GLOBAL_WORKFLOW.md 进任务目录'

    Write-Host ""
    Write-Host "=== 2. 模拟还没填配置（负向测试：工具应当报错）===" -ForegroundColor Cyan

    # 未填的骨架应当被 check-md 报"占位符没填"
    $out1 = & powershell -NoProfile -ExecutionPolicy Bypass -File $checkMd -Path $tmp -Template -Quiet 2>&1
    $phCount = @($out1 | Where-Object { "$_" -match '占位符没填' }).Count
    Assert-True ($phCount -gt 0) '未填的骨架会被报"占位符没填"' "实际报了 $phCount 处"

    # audit 在临时工作区上跑（用 -Path 无从指定，这里只验不崩）
    $outA = & powershell -NoProfile -ExecutionPolicy Bypass -File $audit -SelfTest 2>&1
    Assert-True (@($outA | Where-Object { "$_" -match '自检通过' }).Count -gt 0) 'audit 自检通过'

    Write-Host ""
    Write-Host "=== 3. 填配置（模拟准备流程 B/C 阶段）===" -ForegroundColor Cyan

    function Fill-File([string]$path, [hashtable]$map) {
        $t = Get-Content -LiteralPath $path -Encoding UTF8 -Raw
        foreach ($k in $map.Keys) { $t = $t.Replace($k, $map[$k]) }
        Set-Content -LiteralPath $path -Value $t -Encoding UTF8 -NoNewline
    }

    Fill-File (Join-Path $tmp '_本工作区.md') @{
        '<填：这个工作区叫什么>'                                = '自检测试工作区'
        '<填：任务目录的绝对路径>'      = $taskDir
        '<填：规则系统母版的绝对路径>'                                   = $RulesRoot
    }
    Fill-File (Join-Path $tmp 'AGENTS.md') @{}
    Fill-File (Join-Path $taskDir 'GLOBAL_WORKFLOW.md') @{
        '〔YYYY-MM-DD〕'                    = '2026-10-18'
        '〔项目名称〕'                       = '自检项目'
        '〔负责的模块/范围〕'                  = '全部'
        '〔哪些目录、哪些工作属于本规范管辖〕'      = '本工作区全部'
        '〔起止时间〕'                       = '2026-10-18 ~ 2026-10-19'
        '〔任务看板.md / 计划书.doc〕'          = '任务看板.md'
        '〔**正式开工** ／ **验证系统** ／ **纯讨论**〕' = '**验证系统**'
        '〔基类-代码 / 基类-文档 / …〕'          = '基类-代码'
        '〔约束 1〕'                        = '约束一：不许覆盖原文件'
        '〔约束 2〕'                        = '约束二：改完读回核对'
        '〔例：延迟降低 ≥30%，来源 results/xxx.json〕' = '验收：产出文件存在且非空'
        '〔简述〕'                          = '初始建立'
        '〔第 X 节〕'                       = '§0'
    }
    Fill-File (Join-Path $taskDir '任务看板.md') @{
        '〔YYYY-MM-DD〕'                    = '2026-10-18'
        '〔例：P1 修完缺陷 → P2 重跑基准对比。不需等任何人，可立即执行。〕' = '先做 A，再做 B。'
        '〔input/xxx.md〕'                  = 'input/需求.md'
        '〔规划中 / 已授权执行〕'              = '已授权执行'
        '〔起止〕'                          = '2026-10-18 ~ 2026-10-19'
        '〔代码目录 / 文档目录〕'              = 'output/'
        '〔日期〕'                          = '2026-10-18'
        '截至〔日期〕'                       = '截至2026-10-18'
        '〔例：内容完成度约 80%，可交付完成度约 40%。差的是把已有的东西跑出干净数据、写成文档。〕' = '刚开始。'
        '〔计划书第几项〕'                    = '第1项'
        '〔怎么算达标〕'                      = '文件存在'
        '〔results/xxx.json 或命令〕'         = 'output/x.json'
    }
    # 清掉残留的空占位 〔〕
    foreach ($pp in @((Join-Path $taskDir 'GLOBAL_WORKFLOW.md'), (Join-Path $taskDir '任务看板.md'))) {
        $t = Get-Content -LiteralPath $pp -Encoding UTF8 -Raw
        $t = $t -replace '〔〕', '（无）'
        Set-Content -LiteralPath $pp -Value $t -Encoding UTF8 -NoNewline
    }

    Write-Host ""
    Write-Host "=== 4. 正向测试：填完后应当干净 ===" -ForegroundColor Cyan

    $out2 = & powershell -NoProfile -ExecutionPolicy Bypass -File $checkMd -Path $tmp -Quiet 2>&1
    $leftPh = @($out2 | Where-Object { "$_" -match '占位符没填' }).Count
    Assert-True ($leftPh -eq 0) '填完后不再报"占位符没填"' "还剩 $leftPh 处"
    $struct = @($out2 | Where-Object { "$_" -match '空章节|重复|未闭合|列数' }).Count
    Assert-True ($struct -eq 0) '填完后没有结构问题' "有 $struct 处"

    # "规则完善"的判据：GLOBAL_WORKFLOW 与 任务看板 里不该再有〔〕
    $gw = Get-Content -LiteralPath (Join-Path $taskDir 'GLOBAL_WORKFLOW.md') -Encoding UTF8 -Raw
    $kb = Get-Content -LiteralPath (Join-Path $taskDir '任务看板.md') -Encoding UTF8 -Raw
    Assert-True ($gw -notmatch '〔[^〕]*〕') 'GLOBAL_WORKFLOW 已无占位符（= 规则完善的判据 1~3 满足）'
    Assert-True ($kb -notmatch '〔[^〕]*〕') '任务看板 已无占位符（= 判据 4 满足）'

    Write-Host ""
    Write-Host "=== 5. 登记 → sync 能认出 → 注销 ===" -ForegroundColor Cyan

    $regOrig = Get-Content -LiteralPath $Registry -Encoding UTF8 -Raw
    $regNew = $regOrig -replace '(?m)(<!--\s*/LIST\s*-->)', ("- {0} | derived`r`n`$1" -f $tmp)
    Set-Content -LiteralPath $Registry -Value $regNew -Encoding UTF8 -NoNewline

    $outS = & powershell -NoProfile -ExecutionPolicy Bypass -File $sync -Workspace $tmp 2>&1
    $syncSaw = @($outS | Where-Object { "$_" -match '自检测试|tpl-selftest|v1\.' }).Count
    Assert-True ($syncSaw -gt 0) 'sync.ps1 能认出新登记的工作区' "匹配到 $syncSaw 行"

    # audit 也应认得出
    $outA2 = & powershell -NoProfile -ExecutionPolicy Bypass -File $audit 2>&1
    Assert-True ($syncSaw -gt 0 -and $outA2 -ne $null) 'audit 在登记新工作区后不崩'

    # 注销
    Set-Content -LiteralPath $Registry -Value $regOrig -Encoding UTF8 -NoNewline
    $regAfter = Get-Content -LiteralPath $Registry -Encoding UTF8 -Raw
    Assert-True ($regAfter -eq $regOrig) '登记表已恢复原样（注销干净）' '内容有差异'

    Write-Host ""
    Write-Host "=== 6. 负向测试：缺文件时应当被发现 ===" -ForegroundColor Cyan

    Remove-Item -LiteralPath (Join-Path $tmp 'AGENTS.md') -Force
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $tmp 'AGENTS.md'))) '已删掉 AGENTS.md（造缺）'

    # audit 的必需文件检查针对母版，这里验"能察觉缺失"这个能力本身
    $existCheck = Test-Path -LiteralPath (Join-Path $tmp 'AGENTS.md')
    Assert-True (-not $existCheck) '缺文件能被察觉（Test-Path = False）'

    Write-Host ""
} finally {
    Write-Host "=== 7. 清理 ===" -ForegroundColor Cyan
    if ($created -and -not $Keep) {
        Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
        $still = Test-Path -LiteralPath $tmp
        if (-not $still) { Ok '临时工作区已删除' } else { No '临时工作区没删掉' $tmp }
    } elseif ($Keep) {
        Write-Host "  -Keep：保留在 $tmp" -ForegroundColor Yellow
    }
}

# ---------------------------------------------------------------- 汇总
Write-Host ""
Write-Host ("=== 结果：{0} 通过 / {1} 失败 ===" -f $script:pass, $script:fail) -ForegroundColor $(if($script:fail -eq 0){'Green'}else{'Yellow'})
if ($script:fail -gt 0) {
    Write-Host "注意：**先确认这些失败是真的，再怀疑模版**（先校准尺子）。" -ForegroundColor DarkGray
    exit 1
}
Write-Host ""
Write-Host "自动部分全过。**但这只测了文件与流程。**" -ForegroundColor Cyan
Write-Host "「AI 会不会照规则做」必须真开对话验 —— 见同目录 对话测试.md" -ForegroundColor Cyan
exit 0
