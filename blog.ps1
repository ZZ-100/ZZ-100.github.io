<#
  blog.ps1 - 学术主页管理工具（Hexo + GitHub Actions）

  用法:
    .\blog.ps1 new "标题"       新建主页文章（自动带 academia: true 并打开编辑器）
    .\blog.ps1 page "标题"      新建独立页面
    .\blog.ps1 edit "标题"      打开已有文章编辑
    .\blog.ps1 serve            本地预览 (http://localhost:4000)
    .\blog.ps1 build            本地构建
    .\blog.ps1 publish "说明"   提交并推送（GitHub Actions 会自动部署）
    .\blog.ps1 help             显示帮助

  提示:
    - 系统可能禁止直接运行 .ps1，请用:
        powershell -ExecutionPolicy Bypass -File .\blog.ps1 new "标题"
#>

param()
$ErrorActionPreference = 'Stop'
$Root = $PSScriptRoot
$node = 'C:\Program Files\nodejs\node.exe'

function Show-Help {
    Get-Content -LiteralPath $MyInvocation.MyCommand.Path -TotalCount 1 -ErrorAction SilentlyContinue | Out-Null
    Get-Content -LiteralPath ($MyInvocation.MyCommand.Path) -ErrorAction SilentlyContinue |
        Select-String -Pattern '^  \\\\.\\\\blog' | ForEach-Object { $_.Line }
}

function New-Article {
    param([string]$Title)
    if (-not $Title) { Write-Host '请提供标题: .\blog.ps1 new "标题"' -ForegroundColor Yellow; return }
    $date = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $file = Join-Path $Root "source\_posts\$Title.md"
    if (Test-Path -LiteralPath $file) { Write-Host "已存在: $file" -ForegroundColor Yellow; return }
    $content = @"
---
title: $Title
date: $date
academia: true
---

# $Title
"@
    Set-Content -LiteralPath $file -Value $content -Encoding UTF8
    Write-Host "已创建文章: $file" -ForegroundColor Green
    Start-Process notepad $file
}

function New-Page {
    param([string]$Title)
    if (-not $Title) { Write-Host '请提供标题: .\blog.ps1 page "标题"' -ForegroundColor Yellow; return }
    $dir = Join-Path $Root "source\$Title"
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
    $file = Join-Path $dir 'index.md'
    if (Test-Path -LiteralPath $file) { Write-Host "已存在: $file" -ForegroundColor Yellow; return }
    $content = @"
---
title: $Title
date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
---

# $Title
"@
    Set-Content -LiteralPath $file -Value $content -Encoding UTF8
    Write-Host "已创建页面: $file" -ForegroundColor Green
    Start-Process notepad $file
}

function Edit-Article {
    param([string]$Title)
    if (-not $Title) { Write-Host '请提供标题: .\blog.ps1 edit "标题"' -ForegroundColor Yellow; return }
    $file = Join-Path $Root "source\_posts\$Title.md"
    if (-not (Test-Path -LiteralPath $file)) { Write-Host "找不到文章: $Title" -ForegroundColor Red; return }
    Start-Process notepad $file
}

function Run-Npm {
    param([string]$ScriptName)
    Push-Location $Root
    try {
        npm.cmd run $ScriptName 2>&1
    } finally {
        Pop-Location
    }
}

function Start-Preview {
    Run-Npm 'server'
}

function Build {
    Run-Npm 'clean'
    Run-Npm 'build'
}

function Publish {
    param([string]$Message)
    if (-not $Message) { $Message = '更新内容' }
    Push-Location $Root
    try {
        git add -A
        if (-not $?) { throw 'git add 失败' }
        git commit -m $Message
        if ($LASTEXITCODE -ne 0 -and $LASTEXITCODE -ne 1) { throw 'git commit 失败' }
        git push origin main
        if (-not $?) { throw 'git push 失败' }
        Write-Host '已推送，GitHub Actions 正在自动部署，1-2 分钟生效。' -ForegroundColor Green
    } finally {
        Pop-Location
    }
}

$cmd = $args[0]
switch ($cmd) {
    'new'     { New-Article $args[1] }
    'page'    { New-Page $args[1] }
    'edit'    { Edit-Article $args[1] }
    'serve'   { Start-Preview }
    'build'   { Build }
    'publish' { Publish $args[1] }
    'help'    { Show-Help }
    default   { Show-Help }
}
