<#
.SYNOPSIS
  audit.ps1 —— 规则系统自检：膨胀、重复、断链、下游漂移。

.DESCRIPTION
  查的是**维护健康度**，不是内容对不对：
    - 核心规则行数/字节超线（膨胀）
    - 跨文件重复的规则句（冗杂）
    - 引用的文件是否存在（断链）
    - 核心规则里 §x 交叉引用指向的节是否存在
    - 下游工作区的回执格式是否与母版一致（漂移）
    - 骨架占位符（该填没填）
    - 目录结构里该有的文件都在
    - **宿主自身**：触发面（三时机）/ 交接指针 / 状态文件版本 / 最新 tag（v1.33）

  自检优先：跑 -SelfTest 先证明这个检查器本身是对的。

.PARAMETER SelfTest
  造已知样本验证检查器自己。

.PARAMETER Quiet
  只输出有问题的项。

.EXAMPLE
  .\audit.ps1
  .\audit.ps1 -SelfTest
#>
[CmdletBinding()]
param(
    [switch]$SelfTest,
    [switch]$Quiet
)

$ErrorActionPreference = 'Stop'
$RulesRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent

# 阈值（跟着 01-核心规则.md 开头的硬线走）
$CoreMaxLines = 170
$CoreWarnLines = 150

# 回执字段：**这是唯一机器可读的字段清单**，格式定义在 `骨架\AGENTS.md` §三。
# 加字段时只改这一处 + 骨架 §三，别在文案里写"几行"（那会变成二十处要同步的重复事实）。
$ReceiptFields = @('【已读】', '【依据】', '【理解】', '【缺口】', '【下一步】', '【模式】')

# ---------------------------------------------------------------- 工具函数
function Read-Text([string]$p) {
    if (-not (Test-Path -LiteralPath $p)) { return $null }
    return (Get-Content -LiteralPath $p -Encoding UTF8 -Raw)
}

# 读「已应用工作区」的清单：**两个来源合并**
#   ① 已应用工作区.md          —— 进仓库，只放**示例**（不含真实路径，可公开）
#   ② 已应用工作区.local.md    —— 本地真实清单，**已 gitignore**，不进仓库
# 这样开源版不带个人路径，本地 audit / sync 照常工作。
function Get-RegistryText {
    $a = Read-Text (Join-Path $RulesRoot '已应用工作区.md')
    $b = Read-Text (Join-Path $RulesRoot '已应用工作区.local.md')
    return (@($a, $b) | Where-Object { $_ }) -join "`n"
}

# 抽出"像规则"的句子：非空、非标题、非表格、非代码、非引用、长度≥15
function Get-RuleSentences([string]$text) {
    $out = @()
    $inFence = $false
    foreach ($ln in ($text -split "`n")) {
        $s = $ln.Trim()
        if ($s -match '^```') { $inFence = -not $inFence; continue }
        if ($inFence) { continue }
        if (-not $s) { continue }
        if ($s -match '^#{1,6}\s') { continue }
        if ($s -match '^\|') { continue }
        if ($s -match '^>') { continue }
        if ($s -match '^---+$') { continue }
        $s2 = $s -replace '^[-*]\s+','' -replace '^\d+[.)]\s+',''
        if ($s2 -match '^<!--|-->$') { continue }        # HTML 注释/模板标记，不算规则句
        if ($s2 -match '^=====') { continue }            # 分隔标记
        if ($s2.Length -ge 15) { $out += $s2 }
    }
    return $out
}

# ---------------------------------------------------------------- 检查项
function Invoke-Audit {
    $issues = [System.Collections.Generic.List[string]]::new()
    $warns  = [System.Collections.Generic.List[string]]::new()

    # --- 1. 核心规则膨胀
    $core = Join-Path $RulesRoot '01-核心规则.md'
    $coreText = Read-Text $core
    if (-not $coreText) {
        $issues.Add('核心规则文件不存在：01-核心规则.md')
    } else {
        # 行数：先去掉尾部换行再 split，否则文件末尾那个换行会多算一行
        # （实测：`-split "`n"` 会把 "a`nb`n" 算成 3 段，实际是 2 行）
        $n = ($coreText.TrimEnd("`r","`n") -split "`n").Count
        $b = (Get-Item -LiteralPath $core).Length
        if ($n -gt $CoreMaxLines) {
            $issues.Add("核心规则膨胀：$n 行 > 硬线 $CoreMaxLines 行。**超了就必须搬去基类。**")
        } elseif ($n -gt $CoreWarnLines) {
            $warns.Add("核心规则偏长：$n 行 > 警示线 $CoreWarnLines（硬线 $CoreMaxLines）")
        }
        if (-not $Quiet) { Write-Host ("  核心规则：{0} 行 / {1} 字节" -f $n, $b) -ForegroundColor DarkGray }
    }

    # --- 2. 跨文件重复的规则句
    # 排除两类——它们**故意**和现行规范说同样的话：
    #   ① `更新日志.md`：历史记录，"为什么这么改"会原文复述规范
    #   ② `待更新\` 下的清单：**从核心规则里摘出来给下游看的**，重复是它的本职
    # 拿这两类判冗余，报出来的全是噪音。
    $dupSkip = @('更新日志.md')
    $mdFiles = Get-ChildItem -LiteralPath $RulesRoot -Recurse -File -Filter '*.md' |
               Where-Object { $_.FullName -notmatch '\\\.git\\|\\待更新\\' }
    $seen = @{}
    foreach ($f in $mdFiles) {
        if ($dupSkip -contains $f.Name) { continue }
        $txt = Read-Text $f.FullName
        if (-not $txt) { continue }
        foreach ($s in (Get-RuleSentences $txt)) {
            if (-not $seen.ContainsKey($s)) { $seen[$s] = @() }
            if ($seen[$s] -notcontains $f.Name) { $seen[$s] += $f.Name }
        }
    }
    $dups = @($seen.GetEnumerator() | Where-Object { $_.Value.Count -ge 2 })
    if ($dups.Count -gt 0) {
        foreach ($d in $dups) {
            $warns.Add(("跨文件重复（{0}）：{1}" -f ($d.Value -join '+'), $d.Key.Substring(0,[Math]::Min(50,$d.Key.Length))))
        }
    }

    # --- 3. 引用的文件是否存在
    # 判据：先把**整个工作区**（母版 + 所有已登记工作区）里的文件名收成集合，
    # 再看被引用的**文件名**在不在这个集合里。
    # 为什么不按相对路径解析：规则里大量用"通名"指代（`任务看板.md` 指任一工作区里那个文件），
    # 按路径判会把正常写法全报成断链——那是噪音，不是发现。
    $refPattern = '`([^`\r\n]+\.(?:md|py|ps1|yml|yaml|json))`'
    $missing = @{}
    $wsRoot = Split-Path $RulesRoot -Parent

    # 收集所有已存在的文件名
    $knownNames = @{}
    $scanRoots = @($wsRoot)
    $regTxt = Get-RegistryText
    if ($regTxt) {
        $inL = $false
        foreach ($ln in ($regTxt -split "`n")) {
            if ($ln -match '<!--\s*LIST\s*-->')  { $inL = $true;  continue }
            if ($ln -match '<!--\s*/LIST\s*-->') { $inL = $false; continue }
            if (-not $inL) { continue }
            if ($ln -notmatch '^\s*-\s+\S') { continue }
            $wsPath = ($ln -replace '^\s*-\s+','') -replace '\|\s*(derived|派生)\s*$',''
            $wsPath = $wsPath.Trim().Trim('`')
            if ($wsPath -match '<|〔|\[你的') { continue }
            if (Test-Path -LiteralPath $wsPath) { $scanRoots += $wsPath }
        }
    }
    foreach ($root in ($scanRoots | Select-Object -Unique)) {
        Get-ChildItem -LiteralPath $root -Recurse -File -Force -ErrorAction SilentlyContinue |
            Where-Object { $_.FullName -notmatch '\\\.git\\|\\__pycache__\\' } |
            ForEach-Object { $knownNames[$_.Name.ToLower()] = $true }
    }

    # 模板/示例占位：不是真文件名，跳过
    $placeholderPat = 'YYYY|MMDD|\bNN\b|\bXX\b|\bxxx\b|主题-报告|输入文件|vX\.Y'

    # 整份跳过断链检查的文件：它们是**故意**提到不存在的东西
    #   对话测试.md —— T4 就是考"问一个不存在的文件，它会不会编"
    $skipRefFiles = @('对话测试.md', '外部对标与更新方案.md')   # 研究文档：引用的是**外部仓库路径**，不是本地链接

    foreach ($srcFile in $mdFiles) {
        if ($skipRefFiles -contains $srcFile.Name) { continue }
        $srcTxt = Read-Text $srcFile.FullName
        if (-not $srcTxt) { continue }
        $srcLines = $srcTxt -split "`n"

        foreach ($m in [regex]::Matches($srcTxt, $refPattern)) {
            $ref = $m.Groups[1].Value.Trim()
            if (-not $ref) { continue }
            if ($ref -match '^(https?:|[A-Za-z]:\\)') { continue }     # 外链/绝对
            if ($ref -match '[\*\?<>]') { continue }                    # 通配
            if ($ref -match $placeholderPat) { continue }               # 示例占位

            $base = Split-Path ($ref -replace '/','\') -Leaf
            if (-not $base) { continue }
            if ($knownNames.ContainsKey($base.ToLower())) { continue }

            # 再试：引用可能是"相对某个已登记工作区"的路径（如 `校园外卖平台\spec\x.md`）。
            # 办法：切段后取**首段/前两段/前三段**到每个扫描根下探——那个首段往往是工作区里的子目录。
            $segs = @(($ref -replace '/','\') -split '\\' | Where-Object { $_ })
            $hitByPrefix = $false
            if ($segs.Count -ge 1) {
                $prefixes = @($segs[0])
                if ($segs.Count -ge 2) { $prefixes += ($segs[0..1] -join '\') }
                if ($segs.Count -ge 3) { $prefixes += ($segs[0..2] -join '\') }
                foreach ($sroot in ($scanRoots | Select-Object -Unique)) {
                    foreach ($pre in $prefixes) {
                        if (Test-Path -LiteralPath (Join-Path $sroot $pre)) { $hitByPrefix = $true; break }
                    }
                    if ($hitByPrefix) { break }
                }
            }
            if ($hitByPrefix) { continue }

            # 两类"不是断链"，按**引用所在行**判断：
            #   ① 标了"待建 / 未建 / ⬜"——故意还没建
            #   ② 在举例（"例："）或在说"已被取代 / 归档 / 旧的"
            $lineText = ''
            foreach ($ln in $srcLines) { if ($ln.Contains($ref)) { $lineText = $ln; break } }

            if ($lineText -match '待建|未建|计划建|TODO|⬜') { continue }
            if ($lineText -match '已被.{0,12}取代|归档|旧的|例：|例:') { continue }
            # ③ 泛指举例：那一行在问"有没有 X"，或并列了多个文件名——都是在举例，不是在指本仓库某文件
            if ($lineText -match '有没有|是否含|比如|例如|诸如|如：') { continue }
            if (([regex]::Matches($lineText, '`[^`]*\.(?:md|py|ps1|json|yaml|yml)`')).Count -gt 1) { continue }

            $key = "$($srcFile.Name) → $ref"
            if (-not $missing.ContainsKey($key)) {
                $t = $lineText.Trim()
                $missing[$key] = if ($t) { $t.Substring(0,[Math]::Min(56,$t.Length)) } else { '' }
            }
        }
    }
    foreach ($k in $missing.Keys) { $warns.Add(("断链：{0}　〔{1}〕" -f $k, $missing[$k])) }

    # --- 4. 核心规则的 §x 交叉引用
    if ($coreText) {
        $sectionNames = @{}
        # 已定义的节号（## 0. / ## 〇、/ ## 1. 都收；§2.1 这类子节归到主号 §2）
        foreach ($m in [regex]::Matches($coreText, '(?m)^##\s*([0-9〇零一二三四五六七八九]+)[、.]')) {
            $sectionNames[$m.Groups[1].Value] = $true
        }
        $sectionNames['0'] = $true
        $sectionNames['〇'] = $true

        # 引用：§x 或 §x.y；**排除 §零不是模板例句 这类普通文字**
        # 判据：§ 后面必须紧跟单个数字/汉字，且后面不能跟一个很长的中文串（那是在说"§零不是…"）
        foreach ($m in [regex]::Matches($coreText, '§([0-9〇])(?:\.\d+)?(?![\u4e00-\u9fa5]{2,})')) {
            $sec = $m.Groups[1].Value
            if (-not $sectionNames.ContainsKey($sec)) {
                $warns.Add("核心规则里引用了不存在的节：§$sec")
            }
        }
    }

    # --- 4b. 骨架的不变式 应当和核心规则 §〇 覆盖同一批
    # 判据：数**编号符号**（①-⑧）出现了几个，不是数表格行——
    # 核心规则为省行把 "③④⑤⑥" 挤在一行，数行会误判。
    if ($coreText) {
        $circNums = [char[]]'①②③④⑤⑥⑦⑧'
        $coreSet = @{}
        foreach ($ch in $circNums) { if ($coreText.Contains($ch)) { $coreSet["$ch"] = $true } }

        $skelAgents = Join-Path (Join-Path $RulesRoot '骨架') 'AGENTS.md'
        $skelText = Read-Text $skelAgents
        if ($skelText -and $coreSet.Count -gt 0) {
            $skelSet = @{}
            foreach ($ch in $circNums) { if ($skelText.Contains($ch)) { $skelSet["$ch"] = $true } }

            $onlyCore = @($coreSet.Keys | Where-Object { -not $skelSet.ContainsKey($_) })
            $onlySkel = @($skelSet.Keys | Where-Object { -not $coreSet.ContainsKey($_) })
            if ($onlyCore.Count -gt 0) {
                $warns.Add("不变式没同步到骨架：核心有 $($onlyCore -join '')，骨架 AGENTS.md 没有——**改了一处忘了另一处**")
            }
            if ($onlySkel.Count -gt 0) {
                $warns.Add("骨架有核心没有的不变式编号：$($onlySkel -join '')（编号对不上）")
            }
            if ($onlyCore.Count -eq 0 -and $onlySkel.Count -eq 0 -and -not $Quiet) {
                Write-Host ("  不变式：核心与骨架覆盖同一批（{0} 个）" -f $coreSet.Count) -ForegroundColor DarkGray
            }
        }
    }

    # --- 4c. 更新机制的自洽性
    # 查三件事：版本.txt 和更新日志最新版一致；当前版本的清单在不在；下游有没有落后
    $verFile = Join-Path $RulesRoot '版本.txt'
    $curVer  = if (Test-Path -LiteralPath $verFile) { (Read-Text $verFile).Trim() } else { $null }
    if (-not $curVer) {
        $issues.Add('缺少 `版本.txt`——更新机制的核心信号，下游靠它判断有没有更新')
    } else {
        # 和更新日志的最新版比
        $logTxt = Read-Text (Join-Path $RulesRoot '更新日志.md')
        $logVer = $null
        if ($logTxt) {
            $mm = [regex]::Match($logTxt, '(?m)^## 版本 (v\d+(?:\.\d+)*)')
            if ($mm.Success) { $logVer = $mm.Groups[1].Value }
        }
        if ($logVer -and $logVer -ne $curVer) {
            $issues.Add("版本号不一致：版本.txt 是 $curVer，更新日志最新条目是 $logVer——**发了版但没改版本.txt**")
        }

        # 当前版本的变更清单在不在
        $listDir = Join-Path $RulesRoot '待更新'
        if (-not (Test-Path -LiteralPath $listDir)) {
            $issues.Add('缺少 `待更新\` 目录——下游无从知道要改什么')
        } elseif (-not (Test-Path -LiteralPath (Join-Path $listDir "$curVer.md"))) {
            $issues.Add("当前版本 $curVer 的变更清单不在：`待更新\$curVer.md`——**发了版但没写清单**")
        }
    }

    # --- 4d. 宿主体检（v1.33）
    # 来源：宿主此前"只遵守规则、不跑装置"——状态文件和仓库对不上时没人知道。
    # 这四项全机械可查：不靠人记，也就不靠自觉。
    $hostAgents = Join-Path $wsRoot 'AGENTS.md'
    $hostAgentsText = Read-Text $hostAgents
    if (-not $hostAgentsText) {
        $issues.Add('宿主缺 AGENTS.md——触发面不存在，规则等于没写')
    } else {
        $triggers = @('开工','改母版','收尾')
        $missTrig = @($triggers | Where-Object { $hostAgentsText -notmatch [regex]::Escape($_) })
        if ($missTrig.Count -gt 0) {
            $warns.Add(("宿主 AGENTS.md 缺时机标记：{0}——它退化成纯目录了，规则就没有触发点" -f ($missTrig -join '、')))
        } elseif (-not $Quiet) {
            Write-Host '  宿主：三个时机齐全（开工 / 改母版 / 收尾）' -ForegroundColor DarkGray
        }
        # 它自己声明了"≤ 2 KB"：声明了就必须机械可验，否则又是"可选=没有规则"
        $hb = (Get-Item -LiteralPath $hostAgents).Length
        if ($hb -gt 2048) {
            $warns.Add(("宿主 AGENTS.md 有 {0} 字节，超过它自己写的 2 KB 上限——**触发面正在长成规则正文**" -f $hb))
        }
    }

    $progFile = Join-Path $RulesRoot '_当前进展.md'
    $progText = Read-Text $progFile
    if (-not $progText) {
        $issues.Add('宿主缺 `_当前进展.md`——交接指针的唯一载体，没有它新会话接不上')
    } else {
        $anchor = @('上次时间：','停在哪：','接着读：')
        $missA = @($anchor | Where-Object { $progText -notmatch [regex]::Escape($_) })
        if ($missA.Count -gt 0) {
            $issues.Add(("_当前进展.md 交接指针不全：缺 {0}——新会话接不上" -f ($missA -join '、')))
        } elseif (-not $Quiet) {
            Write-Host '  宿主：交接指针三段齐全' -ForegroundColor DarkGray
        }
        # 状态文件里的"最近版本"表首行，应当等于 版本.txt（过期的状态文件比没有更误导）
        if ($curVer) {
            $mv = [regex]::Match($progText, '(?m)^\|\s*\*\*(v\d+(?:\.\d+)*)\*\*\s*\|')
            if ($mv.Success -and $mv.Groups[1].Value -ne $curVer) {
                $warns.Add(("_当前进展.md 的最近版本写的是 {0}，版本.txt 是 {1}——**状态文件过期**" -f $mv.Groups[1].Value, $curVer))
            }
        }
    }

    # 最新 tag 与 版本.txt 一致（忘了打 tag = 下游按 tag 找版本时会找不到）
    if ($curVer) {
        try {
            $tag = (& git -C $RulesRoot describe --tags --abbrev=0 2>$null)
            if ($LASTEXITCODE -eq 0 -and $tag) {
                $tag = "$tag".Trim()
                if ($tag -ne $curVer) {
                    $warns.Add("最新 tag 是 $tag，版本.txt 是 $curVer——**发了版没打 tag**（或打了 tag 没改版本.txt）")
                } elseif (-not $Quiet) {
                    Write-Host ("  宿主：tag 与版本一致（{0}）" -f $curVer) -ForegroundColor DarkGray
                }
            }
        } catch { }
    }

    # --- 5. 下游回执格式漂移
    $regText = Get-RegistryText
    $regPaths = @()          # 收集登记表里的路径，给 5b 的孤儿检测用
    if ($regText) {
        $inList = $false
        foreach ($line in ($regText -split "`n")) {
            if ($line -match '<!--\s*LIST\s*-->')  { $inList = $true;  continue }
            if ($line -match '<!--\s*/LIST\s*-->') { $inList = $false; continue }
            if (-not $inList) { continue }
            if ($line -notmatch '^\s*-\s+\S') { continue }
            $p = ($line -replace '^\s*-\s+','') -replace '\|\s*(derived|派生)\s*$',''
            $p = $p.Trim().Trim('`')
            # 占位符行（示例/待填）不是真路径，跳过——开源版里必有这种行
            if ($p -match '<|〔|\[你的') { continue }
            $regPaths += $p.TrimEnd('\')
            # 登记了但目录不存在 —— **必须报，不能静默跳过**
            # 否则"路径写错了"和"根本没登记"在审计结果里外观完全一样（同一个病，第 4 个变种）
            if (-not (Test-Path -LiteralPath $p)) {
                $warns.Add("登记表有这一行但目录不存在：$p —— **写错路径和没登记外观一样**，修掉或删掉这一行")
                continue
            }

            $wsAgents = Join-Path $p 'AGENTS.md'
            $wsText = Read-Text $wsAgents
            if (-not $wsText) {
                $warns.Add("下游 $(Split-Path $p -Leaf)：没有 AGENTS.md，无法核对回执格式")
                continue
            }
            $miss = @($ReceiptFields | Where-Object { $wsText -notmatch [regex]::Escape($_) })
            if ($miss.Count -gt 0) {
                $issues.Add(("下游漂移：{0} 的 AGENTS.md 缺回执字段 {1}" -f (Split-Path $p -Leaf), ($miss -join ' ')))
            } elseif (-not $Quiet) {
                Write-Host ("  下游 {0}：回执字段齐全" -f (Split-Path $p -Leaf)) -ForegroundColor DarkGray
            }

            # 不变式数量（v1.34）：派生型最容易在"自己重写规则"时丢掉某一条，
            # 而丢一条就等于丢一个验收点。数编号符号是最便宜的机械判据。
            # 实测（v1.34 上线时）：4 个下游里 3 个不足 6 条。
            $invMarks = @('①','②','③','④','⑤','⑥') | Where-Object { $wsText.Contains($_) }
            if ($wsText -match '不变式' -and $invMarks.Count -lt 6) {
                $warns.Add(("下游 {0}：AGENTS.md 的不变式只有 {1} 个（{2}）——**派生型也要守 6 条**（v1.15 起）" -f (Split-Path $p -Leaf), $invMarks.Count, ($invMarks -join '')))
            }

            # 更新机制：母版路径有没有记、版本落不落后
            $wsSync = Read-Text (Join-Path $p '_同步.md')
            $leaf = Split-Path $p -Leaf
            if (-not $wsSync) {
                $warns.Add("下游 $leaf：没有 _同步.md，开工自检无从下手")
            } else {
                # 注意：这里必须容忍 **母版路径** 这种加粗写法——实测漏过一次，是正则的锅不是文件的锅
                $mPath = [regex]::Match($wsSync, '母版路径\*{0,2}\s*\|\s*`([^`]+)`')
                if (-not $mPath.Success) {
                    $warns.Add("下游 $leaf：_同步.md 没写「母版路径」，开工自检找不到母版")
                }
                # 版本口径（v1.33）：只认**事实**那一栏「已同步到」。
                # 旧的「模版版本」栏是"目标/接入时"的意思，但实测被预写成新版本 3 次，
                # 结果是"已对齐"和"没对齐"外观一样。现在两栏并存且不一致 → 直接报出来。
                $mSynced = [regex]::Match($wsSync, '已同步到\*{0,2}\s*\|\s*`?(v\d+(?:\.\d+)*)`?')
                $mVer    = [regex]::Match($wsSync, '(?<!已同步到)模版版本\*{0,2}\s*\|\s*`?(v\d+(?:\.\d+)*)`?')
                $mine    = $null
                if ($mSynced.Success) { $mine = $mSynced.Groups[1].Value }
                elseif ($mVer.Success) { $mine = $mVer.Groups[1].Value }
                if ($mSynced.Success -and $mVer.Success -and $mSynced.Groups[1].Value -ne $mVer.Groups[1].Value) {
                    $warns.Add("下游 $leaf：`_同步.md` 里「模版版本」= $($mVer.Groups[1].Value)、「已同步到」= $($mSynced.Groups[1].Value)——**两栏打架**（v1.33 上线时 4 个下游全部命中）。按 v1.33 只留「已同步到」")
                }
                # 老布局（「模版版本」+「同步到哪一步」两栏）的同一个病：目标值被预写进事实栏
                $mProg = [regex]::Match($wsSync, '同步到哪一步\*{0,2}\s*\|[^\r\n]*?(v\d+(?:\.\d+)*)')
                if ($mVer.Success -and $mProg.Success -and $mProg.Groups[1].Value -ne $mVer.Groups[1].Value) {
                    $warns.Add("下游 $leaf：`_同步.md` 的「模版版本」= $($mVer.Groups[1].Value)，但「同步到哪一步」只到 $($mProg.Groups[1].Value)——**两栏打架**（v1.33 上线时 4 个下游全部命中）。以正文为准，按 v1.33 改成单一「已同步到」")
                }
                if ($mine -and $curVer) {
                    if ($mine -ne $curVer) {
                        $warns.Add("下游 $leaf 落后：它记的是 $mine，母版是 $curVer")
                    } elseif (-not $Quiet) {
                        Write-Host ("  下游 {0}：版本 {1}（已对齐）" -f $leaf, $mine) -ForegroundColor DarkGray
                    }
                } elseif (-not $mSynced.Success -and -not $mVer.Success) {
                    $warns.Add("下游 $leaf：`_同步.md` 里没记「已同步到 vX.Y」——开工自检无从判断落后没落后")
                }

                # 部署指纹（来源：Spec Kit 的 install manifest）
                # 它要解决的事：母版推更新前，先知道"下游有没有改过这个文件"——
                #   改过 → 必须先问，不能直接覆盖（那会吃掉下游的合法修改）
                #   没改 → 可以直接按清单推
                # 记录方式：_同步.md 的「部署指纹」表里存**同步当时**的 SHA256 前 16 位。
                #   注意：记的是"同步时的样子"，不是"骨架现在的样子"——
                #   因为派生型工作区本来就该和骨架不同，拿骨架比永远不相等，那是噪音。
                $fpRows = [regex]::Matches($wsSync, '(?m)^\|\s*`?([^`|\r\n]+?\.md)`?\s*\|\s*`?([0-9a-fA-F]{8,64})`?\s*\|')
                if ($fpRows.Count -eq 0) {
                    $warns.Add("下游 $leaf：_同步.md 没记「部署指纹」——母版推更新时无法判断它改没改过文件")
                } else {
                    foreach ($fr in $fpRows) {
                        $relName = $fr.Groups[1].Value.Trim()
                        $wantHash = $fr.Groups[2].Value.Trim().ToLower()
                        $fp = Join-Path $p $relName
                        if (-not (Test-Path -LiteralPath $fp)) {
                            $warns.Add("下游 $leaf：指纹里记了 $relName，但文件不存在")
                            continue
                        }
                        $nowHash = (Get-FileHash -LiteralPath $fp -Algorithm SHA256).Hash.ToLower()
                        if (-not $nowHash.StartsWith($wantHash)) {
                            $warns.Add("下游 $leaf：$relName 在同步后被改过（指纹不符）——**推更新前先看 diff，别直接覆盖**")
                        } elseif (-not $Quiet) {
                            Write-Host ("  下游 {0}：{1} 指纹一致（同步后未改）" -f $leaf, $relName) -ForegroundColor DarkGray
                        }
                    }
                }
            }
        }
    }

    # --- 5b. 登记表孤儿检测（**有界扫描，不扫全盘**）
    # 反方向：磁盘上有下游工作区，但 `已应用工作区.md` 没写它。
    # 这类病出现过 **三次**，每次都靠偶然发现：
    #   ① `软件项目管理` 从没登记  ② `基类-验证.md` 没进清单（v1.22 的孤儿检测抓的）
    #   ③ `科教` 从没登记（看会话记录才发现）
    # 扫描判据：目录里**同时有 `AGENTS.md` 和 `_同步.md`**——这是本模版独有的指纹，
    #           普通项目不会有，所以假阳性极低。只在几个已知父目录的**下一层**找，不递归全盘。
    $scanRoots = @(
        (Split-Path $RulesRoot -Parent),
        (Join-Path $env:USERPROFILE 'Desktop'),
        (Join-Path $env:USERPROFILE 'source\repos')
    )
    $unregistered = @()
    foreach ($root in $scanRoots) {
        if (-not (Test-Path -LiteralPath $root)) { continue }
        Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue | ForEach-Object {
            $cand = $_.FullName.TrimEnd('\')
            # 排除母版自己（`规则系统` 下没有，但 `骨架\` 有 AGENTS.md + _同步.md，必须挡掉）
            if ($cand -like "$($RulesRoot.TrimEnd('\'))*") { return }
            if (-not (Test-Path -LiteralPath (Join-Path $cand 'AGENTS.md'))) { return }
            if (-not (Test-Path -LiteralPath (Join-Path $cand '_同步.md')))  { return }
            $hit = $false
            foreach ($rp in $regPaths) { if ($cand -ieq $rp) { $hit = $true; break } }
            if (-not $hit) { $unregistered += $cand }
        }
    }
    if ($unregistered.Count -gt 0) {
        foreach ($u in ($unregistered | Select-Object -Unique)) {
            $warns.Add("登记表孤儿：$u 有 AGENTS.md + _同步.md（是本模版的下游），但没写进 `已应用工作区.md`——**audit 和 sync 都不会管它**")
        }
    } elseif (-not $Quiet) {
        Write-Host ("  登记表：扫描到 {0} 个已登记下游，无孤儿" -f $regPaths.Count) -ForegroundColor DarkGray
    }

    # --- 6. 骨架占位符（走 check-md）
    # 注意：check-md 会往屏幕打一堆逐行输出，这里必须完全吞掉，否则自检和正式跑都会被噪音污染。
    $checkMd = Join-Path $PSScriptRoot 'check-md.ps1'
    if (Test-Path -LiteralPath $checkMd) {
        $skel = Join-Path $RulesRoot '骨架'
        if (Test-Path -LiteralPath $skel) {
            $phCount = 0
            try {
                # 用单独进程跑，彻底隔离它的控制台输出
                $exe = (Get-Command powershell -ErrorAction SilentlyContinue)
                if ($exe) {
                    $raw = & powershell -NoProfile -ExecutionPolicy Bypass -File $checkMd -Path $skel -Quiet 2>&1
                    $phCount = @($raw | Where-Object { "$_" -match '占位符没填' }).Count
                }
            } catch { $phCount = 0 }
            if ($phCount -gt 0) {
                $warns.Add("骨架有 $phCount 处待填占位符（模板本该留空，属正常；列出来是提醒别误当成品）")
            }
        }
    }

    # --- 5c. 候选池（沉积）积压检查
    # 来源：反哺流程。沉积\ 是"候补经验"的暂存区，判据只有两条：要么升格，要么删。
    # 建了池子没检查 = 第 4 个死索引（前面已经栽过三次：没登记 / 没进清单 / 登记路径写错）。
    $deposit = Join-Path $RulesRoot '沉积'
    if (Test-Path -LiteralPath $deposit) {
        $items = @(Get-ChildItem -LiteralPath $deposit -File -Filter '*.md' | Where-Object { $_.Name -ne 'README.md' })
        if ($items.Count -gt 5) {
            $warns.Add("候选池积压：沉积\ 里有 $($items.Count) 条（阈値 5）——**该消化了：升格或删**")
        } elseif ($items.Count -gt 0) {
            $oldest = ($items | Sort-Object LastWriteTime | Select-Object -First 1)
            $age = [int]((Get-Date) - $oldest.LastWriteTime).TotalDays
            if ($age -gt 90) {
                $warns.Add("候选池积压：沉积\ 里最老的条目已 $age 天没动（$($oldest.Name)）——**升格或删，别无限躺着**")
            } elseif (-not $Quiet) {
                Write-Host ("  候选池：{0} 条，最老 {1} 天" -f $items.Count, $age) -ForegroundColor DarkGray
            }
        }
    }
    # --- 6b. 基类孤儿检测（双向：清单指向的文件要存在；存在的文件要在清单里）
    # 来源：agent-os 的 `index-standards` 命令——它的索引检查两个方向：
    #   New files（有文件没索引）/ Deleted files（索引指向已删文件）。
    # 我们原来只有"断链"那一半（检查 §3），缺"孤儿"这一半——新加基类但忘了写进清单，没人会发现。
    $baseDir = Join-Path $RulesRoot '基类'
    $indexFile = Join-Path $baseDir '_清单.md'
    if ((Test-Path -LiteralPath $baseDir) -and (Test-Path -LiteralPath $indexFile)) {
        $indexTxt = Read-Text $indexFile
        $orphans = @()
        Get-ChildItem -LiteralPath $baseDir -File -Filter '*.md' | Where-Object { $_.Name -ne '_清单.md' } | ForEach-Object {
            if ($indexTxt -notmatch [regex]::Escape($_.Name)) { $orphans += $_.Name }
        }
        if ($orphans.Count -gt 0) {
            $warns.Add(("基类孤儿：{0} 在 `基类\` 里但没进 `_清单.md`——**加了基类忘记登记**" -f ($orphans -join '、')))
        } elseif (-not $Quiet) {
            Write-Host ("  基类：{0} 个文件全部在清单里" -f (@(Get-ChildItem -LiteralPath $baseDir -File -Filter '*.md').Count - 1)) -ForegroundColor DarkGray
        }
    }

    # --- 6c. 工具脚本必须有 UTF-8 BOM（**这个坑反复出现**）
    # 实测：`edit` 工具写 .ps1 会剥掉 BOM；PowerShell 5.1 会按 GBK 读中文 → 报一堆 "Missing ')'"。
    # 人工补 BOM 补了好几次，所以做成机械检查。
    Get-ChildItem -LiteralPath (Join-Path $RulesRoot '基类\工具') -File -Filter '*.ps1' -ErrorAction SilentlyContinue | ForEach-Object {
        $head = [System.IO.File]::ReadAllBytes($_.FullName)
        if ($head.Length -lt 3 -or $head[0] -ne 0xEF -or $head[1] -ne 0xBB -or $head[2] -ne 0xBF) {
            $issues.Add("工具脚本缺 BOM：$($_.Name) —— PowerShell 5.1 会按 GBK 读中文，直接语法崩")
        }
    }
    # --- 7. 必需文件是否都在
    $required = @(
        'README.md','母版立场.md','反哺流程.md','00-准备流程.md','01-核心规则.md','02-模式.md','03-子代理.md',
        '更新日志.md','已应用工作区.md','LICENSE','宿主规程.md','_当前进展.md',
        '骨架\AGENTS.md','骨架\_本工作区.md','骨架\_同步.md','骨架\GLOBAL_WORKFLOW.md','骨架\任务看板.md',
        '基类\_清单.md',
        '基类\工具\check-md.ps1','基类\工具\sync.ps1','基类\工具\会话用量.py'
    )
    foreach ($r in $required) {
        if (-not (Test-Path -LiteralPath (Join-Path $RulesRoot $r))) {
            $issues.Add("缺少必需文件：$r")
        }
    }

    return [pscustomobject]@{ Issues = $issues; Warns = $warns }
}

# ---------------------------------------------------------------- 自检
function Invoke-SelfTest {
    Write-Host "=== 自检：用已知样本验证检查器 ===" -ForegroundColor Cyan
    $pass = 0; $fail = 0
    $tmp = Join-Path $env:TEMP ("audit-selftest-" + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $tmp -Force | Out-Null

    # 样本 1：两句一模一样的规则句 → 应当被"跨文件重复"抓到
    $a = "## 一`n`n- 这是一条足够长的规则句子用来测试重复检测功能`n"
    $b = "## 二`n`n- 这是一条足够长的规则句子用来测试重复检测功能`n"
    [System.IO.File]::WriteAllText((Join-Path $tmp 'a.md'), $a, (New-Object System.Text.UTF8Encoding($false)))
    [System.IO.File]::WriteAllText((Join-Path $tmp 'b.md'), $b, (New-Object System.Text.UTF8Encoding($false)))
    $sa = Get-RuleSentences (Read-Text (Join-Path $tmp 'a.md'))
    $sb = Get-RuleSentences (Read-Text (Join-Path $tmp 'b.md'))
    if ($sa[0] -eq $sb[0]) { $pass++; Write-Host "  OK   规则句提取一致（重复检测的前提成立）" -ForegroundColor Green }
    else { $fail++; Write-Host "  FAIL 规则句提取不一致" -ForegroundColor Red }

    # 样本 2：断链检测
    $missingKey = 'nope-does-not-exist.md'
    if (-not (Test-Path -LiteralPath (Join-Path $RulesRoot $missingKey))) {
        $pass++; Write-Host "  OK   不存在的文件确实判为不存在" -ForegroundColor Green
    } else { $fail++; Write-Host "  FAIL 断链判据失效" -ForegroundColor Red }

    # 样本 3：回执字段检查——**行为校验，不锁字段数**
    # 不写成"Count -eq 6"：那等于把"有几个字段"这个会变的事实又抄了一遍。
    # 改成验机制本身：全写了 → 一个都不该报；一个没写 → 每个都该报。
    $allText = ($ReceiptFields -join "`n")
    $missAll  = @($ReceiptFields | Where-Object { $allText      -notmatch [regex]::Escape($_) })
    $missNone = @($ReceiptFields | Where-Object { '一段普通文字' -notmatch [regex]::Escape($_) })
    if ($missAll.Count -eq 0 -and $missNone.Count -eq $ReceiptFields.Count) {
        $pass++; Write-Host ("  OK   回执字段检查有效（{0} 个字段，全有/全无都测得对）" -f $ReceiptFields.Count) -ForegroundColor Green
    } else { $fail++; Write-Host "  FAIL 回执字段检查失效" -ForegroundColor Red }

    # 样本 4：真跑一次，确认不崩
    try {
        $null = Invoke-Audit
        $pass++; Write-Host "  OK   真跑一轮不崩" -ForegroundColor Green
    } catch {
        $fail++; Write-Host ("  FAIL 真跑崩了：{0}" -f $_.Exception.Message) -ForegroundColor Red
    }

    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue

    Write-Host ""
    if ($fail -eq 0) { Write-Host ("自检通过（{0} 项）。可以信。" -f $pass) -ForegroundColor Green; return 0 }
    Write-Host ("自检失败（{0} 通过 / {1} 失败）。**结果不要拿来下结论。**" -f $pass, $fail) -ForegroundColor Red
    return 1
}

if ($SelfTest) { exit (Invoke-SelfTest) }

# ---------------------------------------------------------------- 跑
Write-Host "=== 规则系统自检 ===" -ForegroundColor Cyan
Write-Host "  根目录：$RulesRoot" -ForegroundColor DarkGray
$r = Invoke-Audit

Write-Host ""
if ($r.Issues.Count -eq 0) {
    Write-Host "硬性问题：0 项" -ForegroundColor Green
} else {
    Write-Host "硬性问题：$($r.Issues.Count) 项" -ForegroundColor Red
    foreach ($i in $r.Issues) { Write-Host ("  ! {0}" -f $i) -ForegroundColor Red }
}

if ($r.Warns.Count -eq 0) {
    Write-Host "提示：0 项" -ForegroundColor Green
} else {
    Write-Host "提示：$($r.Warns.Count) 项" -ForegroundColor Yellow
    foreach ($w in $r.Warns) { Write-Host ("  - {0}" -f $w) -ForegroundColor Gray }
}

Write-Host ""
if ($r.Issues.Count -eq 0) { exit 0 } else { exit 1 }
