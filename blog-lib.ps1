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
    param([string]$DocxPath, [string]$Title, [string]$Root, [switch]$Force, [string]$TargetFile)
    $date = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $safeTitle = $Title -replace '[\\/:*?"<>|]', '_'
    if ($TargetFile) {
        $file = $TargetFile
    } else {
        $file = Join-Path $Root "source\_posts\$safeTitle.md"
    }
    if (Test-Path -LiteralPath $file) {
        if (-not $Force) {
            throw "文章已存在: $safeTitle"
        }
        # 覆盖更新时保留原 front-matter 的 title/date（保证 URL 不变）
        $old = Get-Content -LiteralPath $file -Raw -Encoding UTF8
        if ($old -match '(?s)^---\r?\n(.*?)\r?\n---') {
            $oldFm = $Matches[1]
            if ($oldFm -match '(?m)^title:\s*(.+)$') { $Title = $Matches[1].Trim() }
            if ($oldFm -match '(?m)^date:\s*(.+)$') { $date = $Matches[1].Trim() }
        }
        if (-not $TargetFile) {
            $safeTitle = $Title -replace '[\\/:*?"<>|]', '_'
            $file = Join-Path $Root "source\_posts\$safeTitle.md"
        }
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

# ---------- 应用学术风格到 Word 文档 ----------
function Apply-AcademicStyles {
    param($Word, $Doc)
    # 页边距：上下 2.54cm，左右 3.17cm
    $Doc.PageSetup.TopMargin = $Word.CentimetersToPoints(2.54)
    $Doc.PageSetup.BottomMargin = $Word.CentimetersToPoints(2.54)
    $Doc.PageSetup.LeftMargin = $Word.CentimetersToPoints(3.17)
    $Doc.PageSetup.RightMargin = $Word.CentimetersToPoints(3.17)

    # 用 WdBuiltinStyle 数字 ID 取内置样式，避免中文版按名称查找失败
    # wdStyleNormal=-1, wdStyleHeading1=-2, wdStyleHeading2=-3, wdStyleListBullet=-19
    # Normal 正文：宋体小四(12pt) + Times New Roman 西文，1.5倍行距，首行缩进2字符，段后6pt
    $normal = $Doc.Styles.Item(-1)
    $normal.Font.NameFarEast = '宋体'
    $normal.Font.NameAscii = 'Times New Roman'
    $normal.Font.NameOther = 'Times New Roman'
    $normal.Font.Size = 12
    $normal.ParagraphFormat.LineSpacingRule = 1
    $normal.ParagraphFormat.FirstLineIndent = $Word.CentimetersToPoints(0.85)
    $normal.ParagraphFormat.SpaceAfter = 6

    # Heading 1：黑体小三(15pt)加粗、居中、无缩进、段后12pt
    $h1 = $Doc.Styles.Item(-2)
    $h1.Font.NameFarEast = '黑体'
    $h1.Font.NameAscii = 'Times New Roman'
    $h1.Font.Size = 15
    $h1.Font.Bold = $true
    $h1.ParagraphFormat.Alignment = 1
    $h1.ParagraphFormat.FirstLineIndent = 0
    $h1.ParagraphFormat.SpaceBefore = 0
    $h1.ParagraphFormat.SpaceAfter = 12

    # Heading 2：黑体四号(14pt)加粗、左对齐、段前12pt
    $h2 = $Doc.Styles.Item(-3)
    $h2.Font.NameFarEast = '黑体'
    $h2.Font.NameAscii = 'Times New Roman'
    $h2.Font.Size = 14
    $h2.Font.Bold = $true
    $h2.ParagraphFormat.Alignment = 0
    $h2.ParagraphFormat.FirstLineIndent = 0
    $h2.ParagraphFormat.SpaceBefore = 12
    $h2.ParagraphFormat.SpaceAfter = 6
}

# ---------- 新建 Word 草稿（自动配置学术风格） ----------
function New-WordDraft {
    param([string]$Title, [string]$SavePath)
    $word = New-Object -ComObject Word.Application
    try {
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $doc = $word.Documents.Add()
        Apply-AcademicStyles -Word $word -Doc $doc

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

# ---------- 将 md 文章转换为带样式的 Word 文档 ----------
# 在文档末尾追加一个段落（Selection 方式）
# 注意：EndKey(6) 只把光标移到最后一个段落符之前（行尾），需先 TypeParagraph
# 新建段落再输入；样式必须显式设置，否则新段落会继承前一段样式。
function Add-WordPara {
    param([object]$Word, [object]$Doc, [string]$Text, [int]$Style)
    $Doc.Activate()
    $sel = $Word.Selection
    if (-not $sel) { throw '无法获取 Word Selection' }
    $sel.EndKey(6)          # wdStory：光标移到文档末尾（最后一个段落行尾）
    $sel.TypeParagraph()    # 新建段落，光标落于新段落开头
    $sel.TypeText($Text)
    $p = $sel.Paragraphs.Item(1)
    $p.Style = $Style       # 显式设置样式（-1=Normal，避免继承前段样式）
    return $p.Range
}

function Convert-MdToWord {
    param([string]$MdPath, [string]$DocxPath)
    $raw = Get-Content -LiteralPath $MdPath -Raw -Encoding UTF8
    $title = [System.IO.Path]::GetFileNameWithoutExtension($MdPath)
    $body = $raw
    if ($raw -match '(?s)^---\r?\n(.*?)\r?\n---\r?\n?(.*)$') {
        $fm = $Matches[1]
        $body = $Matches[2]
        if ($fm -match '(?m)^title:\s*(.+)$') { $title = $Matches[1].Trim() }
    }

    $word = New-Object -ComObject Word.Application
    try {
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $doc = $word.Documents.Add()
        $null = $word.Selection   # 预初始化 Selection（Visible=false 时首次访问可能为 null）
        Apply-AcademicStyles -Word $word -Doc $doc

        # 标题
        $p = $doc.Paragraphs.Item(1)
        $p.Range.Text = $title
        $p.Style = -2
        try {
            $prop = $doc.BuiltInDocumentProperties.Item('Title')
            if ($prop) { $prop.Value = $title }
        } catch { }

        # 正文逐行转换（# 标题 / **加粗** / - 列表 / 段落）
        foreach ($line in ($body -split "`r?`n")) {
            $trimmed = $line.Trim()
            if (-not $trimmed) { continue }
            if ($trimmed -match '^#{2,6}\s+(.*)') {
                Add-WordPara -Word $word -Doc $doc -Text $Matches[1] -Style -3
                continue
            }
            if ($trimmed -match '^#\s+(.*)') {
                Add-WordPara -Word $word -Doc $doc -Text $Matches[1] -Style -2
                continue
            }
            if ($trimmed -match '^[-*]\s+(.*)') {
                Add-WordPara -Word $word -Doc $doc -Text $Matches[1] -Style -19   # List Bullet
                continue
            }
            # 普通段落 + 加粗
            $clean = $line -replace '\*\*(.+?)\*\*', '$1'
            $r = Add-WordPara -Word $word -Doc $doc -Text $clean -Style -1
            $offset = 0
            foreach ($m in [regex]::Matches($line, '\*\*(.+?)\*\*')) {
                $start = $r.Start + ($m.Index - $offset)
                $r2 = $r.Duplicate
                $r2.Start = $start
                $r2.End = $start + $m.Groups[1].Length
                $r2.Font.Bold = $true
                $offset += 4
            }
        }

        $doc.SaveAs2($DocxPath, 12)
        $doc.Close($false)
        return $true
    } finally {
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
}

# ---------- 将页面 Word 文档转换回页面 md（页面往返编辑用） ----------
# 按段落样式重建 markdown：Heading2→##、List Bullet→- 行、Normal→正文；
# 首段 Heading1（页面标题）保留原 front-matter title 不覆盖。
function Convert-WordToPageMd {
    param([string]$DocxPath, [string]$PageMd)
    $old = Get-Content -LiteralPath $PageMd -Raw -Encoding UTF8
    $fm = ''
    if ($old -match '(?s)^---\r?\n(.*?)\r?\n---') { $fm = $Matches[1].TrimEnd() }

    $word = New-Object -ComObject Word.Application
    try {
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $doc = $word.Documents.Open($DocxPath, $false, $true)
        $styleH1 = $doc.Styles.Item(-2).NameLocal
        $styleH2 = $doc.Styles.Item(-3).NameLocal
        $styleBullet = $doc.Styles.Item(-19).NameLocal

        $lines = New-Object System.Collections.Generic.List[string]
        $skipTitle = $true
        foreach ($p in $doc.Paragraphs) {
            $name = $p.Style.NameLocal
            $text = ($p.Range.Text -replace '[\r\n\x07]', '').Trim()
            if (-not $text) { continue }
            $bold = ($p.Range.Bold -eq -1)
            if ($name -eq $styleH1) {
                if ($skipTitle) { $skipTitle = $false; continue }
                $lines.Add("# $text")
            } elseif ($name -eq $styleH2) {
                $lines.Add("## $text")
            } elseif ($name -eq $styleBullet) {
                if ($bold) { $lines.Add("- **$text**") } else { $lines.Add("- $text") }
            } else {
                if ($bold) { $lines.Add("**$text**") } else { $lines.Add($text) }
            }
        }
        $doc.Close($false)

        # 组装 md：列表行连续、其余行之间空行（保证 markdown 段落/列表正确解析）
        $body = New-Object System.Collections.Generic.List[string]
        $prev = ''
        foreach ($line in $lines) {
            $cur = if ($line -match '^[-*] ') { 'list' } elseif ($line -match '^#{1,6} ') { 'head' } else { 'para' }
            if ($body.Count) {
                if ($prev -eq 'list' -and $cur -eq 'list') { $body.Add($line) }
                else { $body.Add(''); $body.Add($line) }
            } else { $body.Add($line) }
            $prev = $cur
        }

        $content = "---`n$fm`n---`n`n" + ($body -join "`n") + "`n"
        Set-Content -LiteralPath $PageMd -Value $content -Encoding UTF8
        return $PageMd
    } finally {
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
}

# ---------- 检测 .word_edits 中比对应 md 新的 Word 文档（未导入的修改） ----------
function Sync-WordEdits {
    param([string]$Root)
    $stale = @()
    $editDir = Join-Path $Root '.word_edits'
    if (-not (Test-Path -LiteralPath $editDir)) { return $stale }
    foreach ($d in (Get-ChildItem -LiteralPath $editDir -Filter '*.docx' -ErrorAction SilentlyContinue)) {
        $base = [System.IO.Path]::GetFileNameWithoutExtension($d.Name)
        $pageMd = Join-Path $Root "source\$base\index.md"
        $postMd = Join-Path $Root "source\_posts\$base.md"
        if (Test-Path -LiteralPath $pageMd) { $target = $pageMd }
        elseif (Test-Path -LiteralPath $postMd) { $target = $postMd }
        else { continue }
        if ($d.LastWriteTime -gt (Get-Item -LiteralPath $target).LastWriteTime) { $stale += $base }
    }
    return $stale
}

# ---------- 更换 CV（复制 PDF 并启用 CV 下载） ----------
function Set-ThemeCV {
    param([string]$PdfPath, [string]$Root)
    $dir = Join-Path $Root 'source\attaches'
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
    $dest = Join-Path $dir 'CV.pdf'
    Copy-Item -LiteralPath $PdfPath -Destination $dest -Force

    # 启用主题配置中的 cv_dl
    $cfg = Join-Path $Root 'themes\Academia\_config.yml'
    $text = [System.IO.File]::ReadAllText($cfg, [System.Text.Encoding]::UTF8)
    if ($text -match 'cv_dl:\r?\n\s+enable: false') {
        $text = $text -replace 'cv_dl:\r?\n\s+enable: false', "cv_dl:`n  enable: true"
        [System.IO.File]::WriteAllText($cfg, $text, (New-Object System.Text.UTF8Encoding($false)))
    }
    return $dest
}

# ---------- 更新作者名与简介 ----------
function Set-ThemeProfile {
    param([string]$Author, [string]$Bio, [string]$Root)
    $cfg = Join-Path $Root 'themes\Academia\_config.yml'
    $text = [System.IO.File]::ReadAllText($cfg, [System.Text.Encoding]::UTF8)
    $text = $text -replace "(?m)^author:\s*'[^']*'", "author: '$Author'"
    $text = $text -replace "(?m)^author_bio:\s*'[^']*'", "author_bio: '$Bio'"
    [System.IO.File]::WriteAllText($cfg, $text, (New-Object System.Text.UTF8Encoding($false)))
}

# ---------- 读取作者名与简介 ----------
function Get-ThemeProfile {
    param([string]$Root)
    $cfg = Join-Path $Root 'themes\Academia\_config.yml'
    $text = [System.IO.File]::ReadAllText($cfg, [System.Text.Encoding]::UTF8)
    $author = ''
    $bio = ''
    if ($text -match "(?m)^author:\s*'([^']*)'") { $author = $Matches[1] }
    if ($text -match "(?m)^author_bio:\s*'([^']*)'") { $bio = $Matches[1] }
    return @{ Author = $author; Bio = $bio }
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





