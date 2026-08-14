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
    # wdStyleNormal=-1, wdStyleHeading1=-2, wdStyleHeading2=-3, wdStyleHeading3=-4, wdStyleListBullet=-19
    # 排版对齐网页主题 Academia：Times New Roman + 仿宋、深灰 #494e52、1.5倍行距、无首行缩进
    $normal = $Doc.Styles.Item(-1)
    $normal.Font.NameFarEast = '仿宋'
    $normal.Font.NameAscii = 'Times New Roman'
    $normal.Font.NameOther = 'Times New Roman'
    $normal.Font.Size = 13.5
    $normal.Font.Color = 0x524E49
    $normal.ParagraphFormat.LineSpacingRule = 1
    $normal.ParagraphFormat.FirstLineIndent = 0
    try { $normal.ParagraphFormat.CharacterUnitFirstLineIndent = 0 } catch { }
    $normal.ParagraphFormat.SpaceAfter = 6

    # Heading 1：主题 h1=30px(22.5pt)、加粗、左对齐
    $h1 = $Doc.Styles.Item(-2)
    $h1.Font.NameFarEast = '仿宋'
    $h1.Font.NameAscii = 'Times New Roman'
    $h1.Font.Size = 22.5
    $h1.Font.Bold = $true
    $h1.Font.Color = 0x524E49
    $h1.ParagraphFormat.Alignment = 0
    $h1.ParagraphFormat.FirstLineIndent = 0
    try { $h1.ParagraphFormat.CharacterUnitFirstLineIndent = 0 } catch { }
    $h1.ParagraphFormat.SpaceBefore = 0
    $h1.ParagraphFormat.SpaceAfter = 12

    # Heading 2：主题 h2=28px(21pt)、加粗、左对齐
    $h2 = $Doc.Styles.Item(-3)
    $h2.Font.NameFarEast = '仿宋'
    $h2.Font.NameAscii = 'Times New Roman'
    $h2.Font.Size = 21
    $h2.Font.Bold = $true
    $h2.Font.Color = 0x524E49
    $h2.ParagraphFormat.Alignment = 0
    $h2.ParagraphFormat.FirstLineIndent = 0
    try { $h2.ParagraphFormat.CharacterUnitFirstLineIndent = 0 } catch { }
    $h2.ParagraphFormat.SpaceBefore = 12
    $h2.ParagraphFormat.SpaceAfter = 6

    # Heading 3：主题 h3=26px(19.5pt)、加粗、左对齐
    $h3 = $Doc.Styles.Item(-4)
    $h3.Font.NameFarEast = '仿宋'
    $h3.Font.NameAscii = 'Times New Roman'
    $h3.Font.Size = 19.5
    $h3.Font.Bold = $true
    $h3.Font.Color = 0x524E49
    $h3.ParagraphFormat.Alignment = 0
    $h3.ParagraphFormat.FirstLineIndent = 0
    try { $h3.ParagraphFormat.CharacterUnitFirstLineIndent = 0 } catch { }
    $h3.ParagraphFormat.SpaceBefore = 12
    $h3.ParagraphFormat.SpaceAfter = 6

    # List Bullet：此 Word 的 -19 为带大缩进的列表样式（leftChars=1600），清零更紧凑
    $bullet = $Doc.Styles.Item(-19)
    try { $bullet.ParagraphFormat.CharacterUnitLeftIndent = 0 } catch { }
    try { $bullet.ParagraphFormat.CharacterUnitFirstLineIndent = 0 } catch { }
    $bullet.ParagraphFormat.LeftIndent = 0
    $bullet.ParagraphFormat.FirstLineIndent = 0
    $bullet.ParagraphFormat.SpaceAfter = 6
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
# COM 无法可靠清除 Word 样式的字符单位首行缩进（PS 5.1 静默失败），保存后直接改 XML：
# 删除 styles.xml 与 document.xml 中所有 w:ind 的 firstLine/firstLineChars 属性（首行缩进）
function Clear-FirstLineIndentXml {
    param([string]$DocxPath)
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::Open($DocxPath, 'Update')
    try {
        foreach ($name in @('word/styles.xml', 'word/document.xml')) {
            $entry = $zip.GetEntry($name)
            if (-not $entry) { continue }
            $sr = New-Object System.IO.StreamReader($entry.Open())
            $xml = $sr.ReadToEnd(); $sr.Dispose()
            $new = [regex]::Replace($xml, '\s*w:firstLineChars="[^"]*"|\s*w:firstLine="[^"]*"', '')
            if ($new -ne $xml) {
                $entry.Delete()
                $ne = $zip.CreateEntry($name)
                $sw = New-Object System.IO.StreamWriter($ne.Open())
                $sw.Write($new); $sw.Dispose()
            }
        }
    } finally { $zip.Dispose() }
}

function Add-WordPara {
    param([object]$Word, [object]$Doc, [string]$Text, [int]$Style)
    # markdown 标记 → Word 格式：***粗斜*** / **粗** / *斜*；
    # 先剥离星号得到纯文本，再按标记在纯文本中的位置设置对应 range 的加粗/斜体
    $clean = New-Object System.Text.StringBuilder
    $marks = New-Object System.Collections.Generic.List[object]
    $i = 0
    while ($i -lt $Text.Length) {
        $rest = $Text.Substring($i)
        if ($rest -match '^\*\*\*(.+?)\*\*\*') {
            $c = $Matches[1]
            $marks.Add([pscustomobject]@{ Start = $clean.Length; Len = $c.Length; Bold = $true; Italic = $true })
            [void]$clean.Append($c)
            $i += 3 + $c.Length + 3
        } elseif ($rest -match '^\*\*(.+?)\*\*') {
            $c = $Matches[1]
            $marks.Add([pscustomobject]@{ Start = $clean.Length; Len = $c.Length; Bold = $true; Italic = $false })
            [void]$clean.Append($c)
            $i += 2 + $c.Length + 2
        } elseif ($rest -match '^\*([^*\s][^*]*?)\*') {
            $c = $Matches[1]
            $marks.Add([pscustomobject]@{ Start = $clean.Length; Len = $c.Length; Bold = $false; Italic = $true })
            [void]$clean.Append($c)
            $i += 1 + $c.Length + 1
        } else {
            [void]$clean.Append($Text[$i])
            $i++
        }
    }
    $plain = $clean.ToString()
    $Doc.Activate()
    $sel = $Word.Selection
    if (-not $sel) { throw '无法获取 Word Selection' }
    $sel.EndKey(6)          # wdStory：光标移到文档末尾（最后一个段落行尾）
    $sel.TypeParagraph()    # 新建段落，光标落于新段落开头
    $sel.TypeText($plain)
    $p = $sel.Paragraphs.Item(1)
    $p.Style = $Style       # 显式设置样式（-1=Normal，避免继承前段样式）
    $rBase = $p.Range
    foreach ($mk in $marks) {
        $r = $rBase.Duplicate
        $r.Start = $rBase.Start + $mk.Start
        $r.End = $rBase.Start + $mk.Start + $mk.Len
        if ($mk.Bold) { $r.Font.Bold = $true }
        if ($mk.Italic) { $r.Font.Italic = $true }
    }
    return $rBase
}

function Get-Marks {
    param([bool]$Bold, [bool]$Italic)
    if ($Bold -and $Italic) { return '***' }
    if ($Bold) { return '**' }
    if ($Italic) { return '*' }
    return ''
}

# 按词读取段内加粗/斜体格式，还原为 markdown 标记；
# 词间空格跟随格式：下一词同格式则空格留在标记内（**Shanghai University**），
# 格式变化处/段尾的空格放在标记外（**bold** and）
function Get-MarkedText {
    param([object]$Range)
    $items = New-Object System.Collections.Generic.List[object]
    foreach ($w in $Range.Words) {
        $wt = $w.Text
        if ($wt -match '^\s') {
            $items.Add([pscustomobject]@{ Space = $true; Text = $wt; Trail = ''; Bold = $false; Italic = $false })
            continue
        }
        $trail = ''
        if ($wt -match '^(\S.*?)(\s+)$') { $wt = $Matches[1]; $trail = $Matches[2] }
        $items.Add([pscustomobject]@{ Space = $false; Text = $wt; Trail = $trail; Bold = ($w.Font.Bold -eq -1); Italic = ($w.Font.Italic -eq -1) })
    }
    $sb = New-Object System.Text.StringBuilder
    $lastB = $null
    $lastI = $null
    $pending = ''
    for ($i = 0; $i -lt $items.Count; $i++) {
        $cur = $items[$i]
        if ($cur.Space) {
            if ($null -ne $lastB -and ($lastB -or $lastI)) { $pending += $cur.Text } else { [void]$sb.Append($cur.Text) }
            continue
        }
        $b = $cur.Bold
        $it = $cur.Italic
        if ($b -ne $lastB -or $it -ne $lastI) {
            if ($null -ne $lastB) { [void]$sb.Append((Get-Marks $lastB $lastI)) }
            if ($pending) { [void]$sb.Append($pending); $pending = '' }
            $marks = Get-Marks $b $it
            if ($marks) {
                $t = $sb.ToString()
                if ($t.Length -gt 0 -and $t[$t.Length - 1] -match '[A-Za-z0-9]') { [void]$sb.Append(' ') }
                [void]$sb.Append($marks)
            }
            $lastB = $b
            $lastI = $it
        }
        [void]$sb.Append($cur.Text)
        if ($cur.Trail) {
            $nextB = $null
            $nextI = $null
            for ($j = $i + 1; $j -lt $items.Count; $j++) {
                if (-not $items[$j].Space) { $nextB = $items[$j].Bold; $nextI = $items[$j].Italic; break }
            }
            if ($null -eq $nextB -or $nextB -ne $b -or $nextI -ne $it) {
                $pending += $cur.Trail
            } else {
                [void]$sb.Append($cur.Trail)
            }
        }
    }
    if ($null -ne $lastB) { [void]$sb.Append((Get-Marks $lastB $lastI)) }
    if ($pending) { [void]$sb.Append($pending) }
    return $sb.ToString().Trim()
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

        # 标题：独立页面（source\xxx\index.md）不写首段标题——页面标题由 front-matter title / 模板渲染，
        # 若写进 Word 会让用户删除后重新生成又出现；仍写入文档属性 Title 以便 Get-WordTitle 识别
        $isPage = $MdPath -match '\\index\.md$'
        if (-not $isPage) {
            $p = $doc.Paragraphs.Item(1)
            $p.Range.Text = $title
            $p.Style = -2
        }
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
            # 普通段落 + 加粗/斜体（Add-WordPara 解析 ** 与 * 标记）
            Add-WordPara -Word $word -Doc $doc -Text $line -Style -1
        }

        # 页面模式不写标题时，Documents.Add 的初始空段会留在第一行 → 删除它
        if ($isPage) {
            $first = $doc.Paragraphs.Item(1)
            if (-not $first.Range.Text.Trim()) { $first.Range.Delete() }
        }

        $doc.SaveAs2($DocxPath, 12)
        $doc.Close($false)
        Clear-FirstLineIndentXml -DocxPath $DocxPath
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
            if ($name -eq $styleH1 -or $name -eq $styleH2) {
                # 标题整段加粗是样式属性，不还原 ** 标记
                $text = ($p.Range.Text -replace '[\r\n\x07]', '').Trim()
                if (-not $text) { continue }
                if ($name -eq $styleH1) {
                    if ($skipTitle) { $skipTitle = $false; continue }
                    $lines.Add("# $text")
                } else {
                    $lines.Add("## $text")
                }
            } else {
                $text = Get-MarkedText -Range $p.Range
                if (-not $text) { continue }
                if ($name -eq $styleBullet) { $lines.Add("- $text") } else { $lines.Add($text) }
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





