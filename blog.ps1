<#
  blog.ps1 - 学术主页管理工具（交互式界面）

  直接运行进入交互式菜单：
    powershell -ExecutionPolicy Bypass -File .\blog.ps1

  也支持命令行参数（用法见脚本头部说明）。
#>
param()

$ErrorActionPreference = 'Stop'
$Root = $PSScriptRoot
. (Join-Path $Root 'blog-lib.ps1')

function New-Article-Interactive {
    Write-Host ''
    $title = Read-Host '请输入文章标题'
    if (-not $title) { Write-Host '已取消' -ForegroundColor Yellow; return }
    try { $file = New-BlankPost -Title $title -Root $Root }
    catch { Write-Host $_.Exception.Message -ForegroundColor Red; Start-Sleep 1; return }
    Write-Host "已创建文章: $file" -ForegroundColor Green
    Write-Host '正在打开编辑器，填完保存关闭即可。' -ForegroundColor Cyan
    Start-Process notepad $file
    Write-Host '按回车返回菜单...' -ForegroundColor Gray
    Read-Host | Out-Null
}

function New-Page-Interactive {
    Write-Host ''
    $title = Read-Host '请输入页面标题'
    if (-not $title) { Write-Host '已取消' -ForegroundColor Yellow; return }
    try { $file = New-BlankPage -Title $title -Root $Root }
    catch { Write-Host $_.Exception.Message -ForegroundColor Red; Start-Sleep 1; return }
    Write-Host "已创建页面: $file" -ForegroundColor Green
    Write-Host '正在打开编辑器...' -ForegroundColor Cyan
    Start-Process notepad $file
    Write-Host '按回车返回菜单...' -ForegroundColor Gray
    Read-Host | Out-Null
}

function Edit-Article-Interactive {
    $posts = Get-ChildItem -LiteralPath (Join-Path $Root 'source\_posts') -Filter '*.md' -ErrorAction SilentlyContinue
    if (-not $posts) {
        Write-Host '还没有任何文章。' -ForegroundColor Yellow
        Start-Sleep 1
        return
    }
    Write-Host ''
    Write-Host '已有文章：' -ForegroundColor Cyan
    for ($i = 0; $i -lt $posts.Count; $i++) {
        Write-Host ("  {0}. {1}" -f ($i + 1), $posts[$i].BaseName)
    }
    Write-Host ('  {0}. 返回' -f ($posts.Count + 1))
    $choice = Read-Host '选择要编辑的编号'
    $n = 0
    if (-not [int]::TryParse($choice, [ref]$n)) { return }
    if ($n -ge 1 -and $n -le $posts.Count) {
        Write-Host '正在打开编辑器...' -ForegroundColor Cyan
        Start-Process notepad $posts[$n - 1].FullName
        Write-Host '按回车返回菜单...' -ForegroundColor Gray
        Read-Host | Out-Null
    }
}

function Start-Preview {
    Write-Host '正在启动本地预览...' -ForegroundColor Cyan
    Start-Process powershell -WorkingDirectory $Root -ArgumentList '-NoExit', '-Command', 'npm.cmd run server'
    Write-Host '预览已在新窗口启动，地址 http://localhost:4000' -ForegroundColor Green
    Write-Host '关闭该窗口即可停止预览。' -ForegroundColor Gray
    Write-Host '按回车返回菜单...' -ForegroundColor Gray
    Read-Host | Out-Null
}

function Build-Local {
    Write-Host '正在本地构建...' -ForegroundColor Cyan
    Push-Location $Root
    try {
        npm.cmd run clean 2>&1 | Out-Null
        npm.cmd run build 2>&1
    } finally {
        Pop-Location
    }
    Write-Host '构建完成（无报错即为正常）。' -ForegroundColor Green
    Write-Host '按回车返回菜单...' -ForegroundColor Gray
    Read-Host | Out-Null
}

function Publish {
    Write-Host ''
    $msg = Read-Host '提交说明（回车使用默认"更新内容"）'
    if (-not $msg) { $msg = '更新内容' }
    try {
        $published = Publish-Git -Message $msg -Root $Root
        if ($published) { Write-Host '已推送，构建与链接检查通过，GitHub Actions 正在自动部署。' -ForegroundColor Green }
        else { Write-Host '没有需要发布的改动。' -ForegroundColor Yellow }
    } catch {
        Write-Host "发布失败: $_" -ForegroundColor Red
    }
    Write-Host '按回车返回菜单...' -ForegroundColor Gray
    Read-Host | Out-Null
}

function Show-Menu {
    Clear-Host
    Write-Host '====================================' -ForegroundColor Cyan
    Write-Host '   学术主页管理工具' -ForegroundColor White
    Write-Host '====================================' -ForegroundColor Cyan
    Write-Host ''
    Write-Host '  1. 新建主页文章' -ForegroundColor Green
    Write-Host '  2. 新建独立页面' -ForegroundColor Green
    Write-Host '  3. 编辑已有文章' -ForegroundColor Green
    Write-Host '  4. 本地预览' -ForegroundColor Green
    Write-Host '  5. 本地构建' -ForegroundColor Green
    Write-Host '  6. 提交并发布' -ForegroundColor Green
    Write-Host '  0. 退出' -ForegroundColor Red
    Write-Host ''
}

function Run-Interactive {
    while ($true) {
        Show-Menu
        $choice = Read-Host '请选择'
        switch ($choice) {
            '1' { New-Article-Interactive }
            '2' { New-Page-Interactive }
            '3' { Edit-Article-Interactive }
            '4' { Start-Preview }
            '5' { Build-Local }
            '6' { Publish }
            '0' { Write-Host '再见！' -ForegroundColor Cyan; return }
            default { Write-Host '无效选项，请重试。' -ForegroundColor Yellow; Start-Sleep 1 }
        }
    }
}

# 命令行兼容
if ($args.Count -gt 0) {
    switch ($args[0]) {
        'new'     { if (-not $args[1]) { Write-Host '用法: .\blog.ps1 new "标题"'; return }; try { $file = New-BlankPost -Title $args[1] -Root $Root; Write-Host "已创建: $file" -ForegroundColor Green; Start-Process notepad $file } catch { Write-Host $_.Exception.Message -ForegroundColor Red } }
        'publish' { Publish }
        default   { Run-Interactive }
    }
    return
}

Run-Interactive
