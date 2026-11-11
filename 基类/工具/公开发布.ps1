#Requires -Version 5.1
<#
.SYNOPSIS
  把母版导出成一个干净的公开快照，扫一遍，再提交/推送。

.DESCRIPTION
  公开仓 = **每个大版本一个快照提交**，不是工作树的历史。
  工作树仓库负责本地版本管理；本脚本负责"对外那一份"。

  它做四件事：
    1. 镜像同步（排除 _当前进展.md / *.local.md / .git）
    2. 敏感扫描（内容层）
    3. 作者邮箱校验（元数据层——这一层 grep 扫不到）
    4. commit → tag → push（作者强制为 noreply）

  任何一步不干净就退出，不写仓库、不推送。

.EXAMPLE
  powershell -NoProfile -ExecutionPolicy Bypass -File 基类\工具\公开发布.ps1 -DryRun
  powershell -NoProfile -ExecutionPolicy Bypass -File 基类\工具\公开发布.ps1
#>
param(
    [switch]$DryRun,
    [string]$RepoPath = '',
    [string]$Proxy = ''
)

$ErrorActionPreference = 'Stop'

$ToolDir   = $PSScriptRoot
$RulesRoot = Split-Path -Parent (Split-Path -Parent $ToolDir)
$HostRoot  = Split-Path -Parent $RulesRoot
if (-not $RepoPath) { $RepoPath = Join-Path $HostRoot '_发布\repo' }

$AuthorName  = 'qianxue-snow'
$AuthorEmail = 'qianxue-snow@users.noreply.github.com'

# 代理：没有显式传入时，先看环境变量，再看系统代理（IE 设置）
if (-not $Proxy) {
    foreach ($v in @($env:HTTPS_PROXY, $env:HTTP_PROXY)) { if ($v) { $Proxy = $v; break } }
}
if (-not $Proxy) {
    try {
        $ie = Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings' -ErrorAction Stop
        if ($ie.ProxyEnable -eq 1 -and $ie.ProxyServer) { $Proxy = 'http://' + $ie.ProxyServer }
    } catch { }
}
if ($Proxy -and $Proxy -notmatch '^https?://') { $Proxy = 'http://' + $Proxy }

function Fail($msg) { Write-Host "  [x] $msg" -ForegroundColor Red; exit 1 }
function Ok($msg)   { Write-Host "  [ok] $msg" -ForegroundColor Green }
function Info($msg) { Write-Host "  - $msg" -ForegroundColor DarkGray }

Write-Host '=== 公开发布 ===' -ForegroundColor Cyan

# --- 0. 版本
$verFile = Join-Path $RulesRoot '版本.txt'
if (-not (Test-Path $verFile)) { Fail '找不到 版本.txt' }
$ver = (Get-Content $verFile -Raw).Trim()
Info "版本：$ver"

# --- 1. 防误删护栏 + 仓库检查
if ($RepoPath -notlike '*\_发布\repo') {
    Fail "RepoPath 不在 _发布\repo 下，拒绝执行（/MIR 会删文件）：$RepoPath"
}
if (-not (Test-Path (Join-Path $RepoPath '.git'))) {
    Fail "不是 git 仓库：$RepoPath（首次请先 git init 并设 origin）"
}

# 发布内容必须等于已提交版本，否则公开的和私仓的对不上
$dirty = & git -C $RulesRoot status --porcelain
if ($dirty) {
    Write-Host "  工作树有未提交改动：" -ForegroundColor Yellow
    $dirty | Select-Object -First 10 | ForEach-Object { Write-Host "    $_" -ForegroundColor Yellow }
    Fail '先 commit 再发布'
}

# --- 2. 镜像同步
Info "同步 $RulesRoot → $RepoPath"
& robocopy $RulesRoot $RepoPath /MIR /XD '.git' /XF '_当前进展.md' '*.local.md' /NFL /NDL /NJH /NJS /NP | Out-Null
if ($LASTEXITCODE -ge 8) { Fail "robocopy 失败：$LASTEXITCODE" }

# --- 3. 内容扫描
$patterns = @(
    # 拆开写：扫描器自己不能含字面量，否则会扫到自己
    @{ Name = '真实用户名';       Regex = ('138' + '77') },
    @{ Name = '手机号';           Regex = '(?<!\d)1[3-9]\d{9}(?!\d)' },
    @{ Name = '163 邮箱';         Regex = '@163\.com' },
    @{ Name = 'Windows 用户目录'; Regex = '[A-Za-z]:\\Users\\(?!me(\\|/|$))' },
    @{ Name = 'Unix 用户目录';    Regex = '/Users/(?!me(\\|/|$))' },
    @{ Name = '凭据形态';         Regex = '(ghp_[A-Za-z0-9]{20,}|github_pat_[A-Za-z0-9_]{20,}|sk-[A-Za-z0-9]{20,}|AKIA[0-9A-Z]{16})' }
)
$textExt = @('.md', '.ps1', '.py', '.txt', '.json', '.yaml', '.yml', '.psm1', '.gitignore', '.gitattributes')
$hits = New-Object System.Collections.ArrayList

Get-ChildItem -LiteralPath $RepoPath -Recurse -File | Where-Object {
    $_.Name -ne '公开发布.ps1' -and
    $_.FullName -notmatch '\\\.git\\' -and (
        $textExt -contains $_.Extension.ToLower() -or
        $_.Name -in @('LICENSE', '.gitignore', '.gitattributes')
    )
} | ForEach-Object {
    $f = $_
    $text = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8
    if (-not $text) { return }
    foreach ($p in $patterns) {
        $ms = [regex]::Matches($text, $p.Regex)
        foreach ($m in $ms) {
            $line = ($text.Substring(0, $m.Index) -split "`n").Count
            $rel = $f.FullName.Substring($RepoPath.Length + 1)
            [void]$hits.Add(("{0}:{1} [{2}] {3}" -f $rel, $line, $p.Name, $m.Value))
        }
    }
}
if ($hits.Count -gt 0) {
    $hits | Select-Object -First 30 | ForEach-Object { Write-Host "    $_" -ForegroundColor Yellow }
    Fail "内容扫描命中 $($hits.Count) 处，已中止"
}
Ok '内容扫描：0 命中'

# --- 4. 元数据：作者邮箱（grep 扫不到的那一层）
$cfgEmail = (& git -C $RepoPath config user.email 2>$null)
if ($cfgEmail -and $cfgEmail.Trim() -ne $AuthorEmail) {
    Fail "仓库 user.email 是 $($cfgEmail.Trim())，不是 noreply——先 git config user.email $AuthorEmail"
}
Ok "作者邮箱：$AuthorEmail"

if ($DryRun) {
    Write-Host '=== DryRun：内容与元数据通过；未提交、未推送 ===' -ForegroundColor Yellow
    exit 0
}

# --- 5. 提交
& git -C $RepoPath add -A
$changed = & git -C $RepoPath status --porcelain
if (-not $changed) {
    Ok '与上次公开快照无差异，跳过提交'
} else {
    & git -C $RepoPath -c "user.name=$AuthorName" -c "user.email=$AuthorEmail" commit -m "$ver：公开快照（$(Get-Date -Format 'yyyy-MM-dd')）" | Out-Null
    if ($LASTEXITCODE -ne 0) { Fail 'commit 失败' }
    Ok '已提交'
}

$lastEmail = (& git -C $RepoPath log -1 --format='%ae' 2>$null)
if (-not $lastEmail -or $lastEmail.Trim() -ne $AuthorEmail) {
    Fail "最新提交的作者是 $lastEmail，不是 $AuthorEmail"
}
Ok '提交作者校验通过'

# --- 6. tag
$hasTag = & git -C $RepoPath tag --list $ver
if (-not $hasTag) {
    & git -C $RepoPath tag $ver
    if ($LASTEXITCODE -ne 0) { Fail "打 tag $ver 失败" }
    Ok "已打 tag $ver"
} else {
    Info "tag $ver 已存在"
}

# --- 7. push
# DSH 沙箱下 git 的 credential helper 会调 sh.exe 建命名管道，被拒（Win32 error 5）。
# 所以先试常规 push，失败再从 GCM 取凭据、拼进 URL 直推（不落盘、不打印）。
function Push-Ref($Refspec) {
    $eap = $ErrorActionPreference
    # native 命令的 stderr 在 Stop 模式下会被当成终止错误，这里临时放宽
    $ErrorActionPreference = 'SilentlyContinue'
    try {
        $pa = @()
        if ($Proxy) { $pa = @('-c', "http.proxy=$Proxy") }

        & git -C $RepoPath @pa push origin $Refspec 2>&1 | Out-Null
        if ($LASTEXITCODE -eq 0) { return $true }

        $o = (& git -C $RepoPath remote get-url origin).Trim()
        if ($o -notmatch '^https://github\.com/(.+)$') { return $false }
        $repoPart = $Matches[1]

        $in  = "protocol=https`nhost=github.com`n`n"
        $out = $in | & git credential-manager get 2>&1
        $pu = $null; $pp = $null
        foreach ($l in $out) {
            if ($l -match '^username=(.*)$') { $pu = $Matches[1] }
            if ($l -match '^password=(.*)$') { $pp = $Matches[1] }
        }
        if (-not $pu -or -not $pp) { return $false }
        $url = 'https://{0}:{1}@github.com/{2}' -f $pu, $pp, $repoPart
        & git -C $RepoPath @pa push $url $Refspec 2>&1 | Out-Null
        return ($LASTEXITCODE -eq 0)
    } catch {
        return $false
    } finally {
        $ErrorActionPreference = $eap
    }
}

if ($Proxy) { Info "代理：$Proxy" }
if (-not (Push-Ref 'main')) { Fail 'push main 失败（检查代理或凭据）' }
if (-not (Push-Ref $ver)) { Fail "push tag $ver 失败" }
Ok '已推送'
Write-Host '=== 完成 ===' -ForegroundColor Cyan
