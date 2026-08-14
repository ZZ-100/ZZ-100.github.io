# blog-lib.ps1 - 学术主页核心函数库（供 GUI/CLI 脚本共用）

$ErrorActionPreference = 'Stop'

function Get-BlogRoot {
    return Split-Path -Parent $MyInvocation.MyCommand.Path
}

# ---------- Word 转 HTML（保留排版） ----------
function Convert-WordToHtml {
    param([string]$DocxPath, [string]$OutHtml)
    $word = New-Object -ComObject Word.Application
    try {
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $doc = $word.Documents.Open($DocxPath, $false, $true)
        # wdFormatHTML = 8, 编码 65001 = UTF-8
        $doc.SaveAs2($OutHtml, 8, $false, '', $false, '', $false, $false, $false, $false, $false, 65001)
        $doc.Close($false)
        return $true
    } finally {
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
}

# ---------- 读取 Word 文档标题 ----------
function Get-WordTitle {
    param([string]$DocxPath)
    $word = New-Object -ComObject Word.Application
    try {
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $doc = $word.Documents.Open($DocxPath, $false, $true)
        $title = ''
        try {
            $title = $doc.BuiltInDocumentProperties.Item('Title').Value
        } catch { }
        if (-not $title) {
            $title = $doc.Paragraphs.Item(1).Range.Text.Trim()
        }
        $doc.Close($false)
        if (-not $title) { $title = [System.IO.Path]::GetFileNameWithoutExtension($DocxPath) }
        return $title.Trim()
    } finally {
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
}

# ---------- 从 Word 生成 markdown 文章（保留排版） ----------
function New-PostFromWord {
    param([string]$DocxPath, [string]$Title, [string]$Root)
    $date = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $safeTitle = $Title -replace '[\\/:*?"<>|]', '_'
    $file = Join-Path $Root "source\_posts\$safeTitle.md"
    if (Test-Path -LiteralPath $file) {
        throw "文章已存在: $safeTitle"
    }

    $tmpHtml = Join-Path $env:TEMP ("hexo_import_" + [System.IO.Path]::GetRandomFileName() + ".htm")
    Convert-WordToHtml -DocxPath $DocxPath -OutHtml $tmpHtml | Out-Null

    # Word 保存的 HTML 实际是系统 ANSI 编码(GBK)，必须按 GBK 解码，否则中文乱码
    $gbk = [System.Text.Encoding]::GetEncoding(936)
    $html = $gbk.GetString([System.IO.File]::ReadAllBytes($tmpHtml))
    # 提取 <style> 块（保存样式排版）
    $styles = ''
    if ($html -match '(?s)<style[^>]*>.*?</style>') { $styles = $Matches[0] }
    # 提取 body 内容
    $body = $html
    if ($html -match '(?s)<body[^>]*>(.*)</body>') { $body = $Matches[1] }
    # 清理 Word 的注释等干扰
    $body = $body -replace '<!\[^>\]*?>', ''
    Remove-Item -LiteralPath $tmpHtml -ErrorAction SilentlyContinue

    $content = @"
---
title: $Title
date: $date
academia: true
---

$styles
$body
"@
    Set-Content -LiteralPath $file -Value $content -Encoding UTF8
    return $file
}

# ---------- 新建空白文章 ----------
function New-BlankPost {
    param([string]$Title, [string]$Root)
    $date = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $safeTitle = $Title -replace '[\\/:*?"<>|]', '_'
    $file = Join-Path $Root "source\_posts\$safeTitle.md"
    if (Test-Path -LiteralPath $file) {
        throw "文章已存在: $safeTitle"
    }
    Set-Content -LiteralPath $file -Value "---`ntitle: $Title`ndate: $date`nacademia: true`n---`n`n# $Title`n" -Encoding UTF8
    return $file
}

# ---------- Git 发布 ----------
function Publish-Git {
    param([string]$Message, [string]$Root)
    if (-not $Message) { $Message = '更新内容' }
    Push-Location $Root
    try {
        git add -A
        if (-not $?) { throw 'git add 失败' }
        git commit -m $Message
        if ($LASTEXITCODE -ne 0 -and $LASTEXITCODE -ne 1) { throw 'git commit 失败' }
        git push origin main
        if (-not $?) { throw 'git push 失败' }
        return $true
    } finally {
        Pop-Location
    }
}

