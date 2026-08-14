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

# ---------- 新建 Word 草稿（自动配置学术风格） ----------
function New-WordDraft {
    param([string]$Title, [string]$SavePath)
    $word = New-Object -ComObject Word.Application
    try {
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $doc = $word.Documents.Add()

        # 页边距：上下 2.54cm，左右 3.17cm
        $doc.PageSetup.TopMargin = $word.CentimetersToPoints(2.54)
        $doc.PageSetup.BottomMargin = $word.CentimetersToPoints(2.54)
        $doc.PageSetup.LeftMargin = $word.CentimetersToPoints(3.17)
        $doc.PageSetup.RightMargin = $word.CentimetersToPoints(3.17)

        # 用 WdBuiltinStyle 数字 ID 取内置样式，避免中文版按名称查找失败
        # wdStyleNormal=-1, wdStyleHeading1=-2, wdStyleHeading2=-3
        # Normal 正文：宋体小四(12pt) + Times New Roman 西文，1.5倍行距，首行缩进2字符，段后6pt
        $normal = $doc.Styles.Item(-1)
        $normal.Font.NameFarEast = '宋体'
        $normal.Font.NameAscii = 'Times New Roman'
        $normal.Font.NameOther = 'Times New Roman'
        $normal.Font.Size = 12
        $normal.ParagraphFormat.LineSpacingRule = 1
        $normal.ParagraphFormat.FirstLineIndent = $word.CentimetersToPoints(0.85)
        $normal.ParagraphFormat.SpaceAfter = 6

        # Heading 1：黑体小三(15pt)加粗、居中、无缩进、段后12pt
        $h1 = $doc.Styles.Item(-2)
        $h1.Font.NameFarEast = '黑体'
        $h1.Font.NameAscii = 'Times New Roman'
        $h1.Font.Size = 15
        $h1.Font.Bold = $true
        $h1.ParagraphFormat.Alignment = 1
        $h1.ParagraphFormat.FirstLineIndent = 0
        $h1.ParagraphFormat.SpaceBefore = 0
        $h1.ParagraphFormat.SpaceAfter = 12

        # Heading 2：黑体四号(14pt)加粗、左对齐、段前12pt
        $h2 = $doc.Styles.Item(-3)
        $h2.Font.NameFarEast = '黑体'
        $h2.Font.NameAscii = 'Times New Roman'
        $h2.Font.Size = 14
        $h2.Font.Bold = $true
        $h2.ParagraphFormat.Alignment = 0
        $h2.ParagraphFormat.FirstLineIndent = 0
        $h2.ParagraphFormat.SpaceBefore = 12
        $h2.ParagraphFormat.SpaceAfter = 6

        # 首行 = 文章标题（Heading 1 居中），并写入文档标题属性
        $r = $doc.Paragraphs.Item(1).Range
        $r.Text = $Title
        $r.Style = -2   # wdStyleHeading1
        try {
            $prop = $doc.BuiltInDocumentProperties.Item('Title')
            if ($prop) { $prop.Value = $Title }
        } catch { }

        # 保存为 docx (wdFormatXMLDocument = 12)
        $doc.SaveAs2($SavePath, 12)
        $doc.Close($false)
        return $true
    } finally {
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
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




