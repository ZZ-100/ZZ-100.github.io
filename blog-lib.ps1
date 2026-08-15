# blog-lib.ps1 - 学术主页核心函数库（供 GUI/CLI 脚本共用）

$ErrorActionPreference = 'Stop'

function Get-BlogRoot {
    return Split-Path -Parent $MyInvocation.MyCommand.Path
}

function ConvertTo-YamlSingleQuoted {
    param([string]$Value)
    return "'" + (($Value -replace '[\r\n]', ' ') -replace "'", "''") + "'"
}

function Get-WordEditBaselinePath {
    param([string]$DocxPath)
    return "$DocxPath.sha256"
}

function Set-WordEditBaseline {
    param([string]$DocxPath)
    if (-not (Test-Path -LiteralPath $DocxPath)) { return }
    $hash = (Get-FileHash -LiteralPath $DocxPath -Algorithm SHA256).Hash
    [System.IO.File]::WriteAllText((Get-WordEditBaselinePath -DocxPath $DocxPath), $hash + "`n", (New-Object System.Text.UTF8Encoding($false)))
}

function Test-WordEditChanged {
    param([string]$DocxPath)
    if (-not (Test-Path -LiteralPath $DocxPath)) { return $false }
    $baselinePath = Get-WordEditBaselinePath -DocxPath $DocxPath
    if (-not (Test-Path -LiteralPath $baselinePath)) { return $true }
    $baseline = (Get-Content -LiteralPath $baselinePath -Raw -Encoding UTF8).Trim()
    $current = (Get-FileHash -LiteralPath $DocxPath -Algorithm SHA256).Hash
    return $baseline -ne $current
}

function Move-ToRecycleBin {
    param([string]$FilePath, [string]$Root)
    if (-not (Test-Path -LiteralPath $FilePath)) { return $null }
    $rootFull = ((Resolve-Path -LiteralPath $Root).Path).TrimEnd('\') + '\'
    $fileFull = (Resolve-Path -LiteralPath $FilePath).Path
    if (-not $fileFull.StartsWith($rootFull, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "只能回收项目目录内的文件: $FilePath"
    }
    $relative = $fileFull.Substring($rootFull.Length)
    $bin = Join-Path $Root ('.trash\' + (Get-Date -Format 'yyyyMMdd-HHmmssfff'))
    $destination = Join-Path $bin $relative
    New-Item -ItemType Directory -Path (Split-Path -Parent $destination) -Force | Out-Null
    Move-Item -LiteralPath $FilePath -Destination $destination -Force
    return $destination
}

function Invoke-HexoBuild {
    param([string]$Root)
    Push-Location $Root
    try {
        $cleanOutput = & npm.cmd run clean 2>&1
        if ($LASTEXITCODE -ne 0) { throw "清理失败:`n$($cleanOutput -join "`n")" }
        $buildOutput = & npm.cmd run build 2>&1
        if ($LASTEXITCODE -ne 0) { throw "构建失败:`n$($buildOutput -join "`n")" }
        return @($buildOutput)
    } finally {
        Pop-Location
    }
}

function Test-PublicLinks {
    param([string]$Root)
    $public = Join-Path $Root 'public'
    if (-not (Test-Path -LiteralPath $public)) { throw '构建目录 public 不存在' }
    $broken = New-Object System.Collections.Generic.List[string]
    foreach ($html in (Get-ChildItem -LiteralPath $public -Recurse -Filter '*.html' -File)) {
        $content = Get-Content -LiteralPath $html.FullName -Raw -Encoding UTF8
        foreach ($match in [regex]::Matches($content, '(?:href|src)=["''](/[^"''#?]*)["'']')) {
            $ref = $match.Groups[1].Value
            $target = Join-Path $public ($ref.TrimStart('/') -replace '/', '\')
            if ($ref.EndsWith('/')) { $target = Join-Path $target 'index.html' }
            if (-not (Test-Path -LiteralPath $target)) {
                $broken.Add("$($html.FullName): $ref")
            }
        }
    }
    if ($broken.Count -gt 0) { throw "发现内部资源或链接不存在:`n$($broken -join "`n")" }
    return $true
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

# ---------- 从 Word 生成 markdown 文章（干净 markdown 重建，与页面路径一致） ----------
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

    # 干净 markdown 重建（与页面路径一致，不嵌入 Word HTML）
    Convert-WordToPageMd -DocxPath $DocxPath -PageMd $file | Out-Null

    # 确保 front-matter 完整：title/date/academia 齐全
    $c = Get-Content -LiteralPath $file -Raw -Encoding UTF8
    $fm = ''
    $body = $c
    if ($c -match '(?s)^---\r?\n(.*?)\r?\n---\r?\n?(.*)$') { $fm = $Matches[1]; $body = $Matches[2] }
    $nl = New-Object System.Collections.Generic.List[string]
    $hasTitle = $hasDate = $hasAcad = $false
    foreach ($ln in ($fm -split "`r?`n")) {
        if ($ln -match '(?i)^title:') { $hasTitle = $true }
        elseif ($ln -match '(?i)^date:') { $hasDate = $true }
        elseif ($ln -match '(?i)^academia:') { $hasAcad = $true }
        $nl.Add($ln)
    }
    if (-not $hasTitle) { $nl.Add("title: $(ConvertTo-YamlSingleQuoted -Value $Title)") }
    if (-not $hasDate) { $nl.Add("date: $date") }
    if (-not $hasAcad) { $nl.Add("academia: true") }
    $content = "---`n" + ($nl -join "`n") + "`n---`n`n" + $body.TrimStart("`r", "`n")
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
    $yamlTitle = ConvertTo-YamlSingleQuoted -Value $Title
    Set-Content -LiteralPath $file -Value "---`ntitle: $yamlTitle`ndate: $date`nacademia: true`n---`n`n# $Title`n" -Encoding UTF8
    return $file
}

# ---------- 新建独立页面 ----------
function New-BlankPage {
    param([string]$Title, [string]$Root)
    $cleanTitle = ($Title -replace '[\r\n]', ' ').Trim()
    $safeTitle = $cleanTitle -replace '[\\/:*?"<>|]', '_'
    if (-not $safeTitle) { throw '页面标题不能为空' }
    $dir = Join-Path $Root "source\$safeTitle"
    $file = Join-Path $dir 'index.md'
    if (Test-Path -LiteralPath $file) { throw "页面已存在: $safeTitle" }
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
    $yamlTitle = ConvertTo-YamlSingleQuoted -Value $cleanTitle
    Set-Content -LiteralPath $file -Value "---`ntitle: $yamlTitle`ndate: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n---`n" -Encoding UTF8
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
    Parse-MdMarks $Text $false $false $clean $marks
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

# 递归解析 markdown 加粗/斜体标记（支持嵌套，如 **粗 *斜* 粗**），
# 结果写入 $Clean（纯文本）与 $Marks（{Start,Len,Bold,Italic} 格式段，相邻同格式已合并）
function Parse-MdMarks {
    param([string]$Text, [bool]$Bold, [bool]$Italic, [System.Text.StringBuilder]$Clean, [System.Collections.Generic.List[object]]$Marks)
    $i = 0
    while ($i -lt $Text.Length) {
        $rest = $Text.Substring($i)
        if ($rest -match '^\*\*\*(.+?)\*\*\*') {
            $c = $Matches[1]
            Parse-MdMarks $c $true $true $Clean $Marks
            $i += 3 + $c.Length + 3
        } elseif ($rest -match '^\*\*(.+?)\*\*') {
            $c = $Matches[1]
            Parse-MdMarks $c $true $Italic $Clean $Marks
            $i += 2 + $c.Length + 2
        } elseif ($rest -match '^\*([^*\s][^*]*?)\*') {
            $c = $Matches[1]
            Parse-MdMarks $c $Bold $true $Clean $Marks
            $i += 1 + $c.Length + 1
        } else {
            $start = $Clean.Length
            [void]$Clean.Append($Text[$i])
            if ($Bold -or $Italic) {
                if ($Marks.Count -gt 0 -and $Marks[$Marks.Count - 1].Bold -eq $Bold -and $Marks[$Marks.Count - 1].Italic -eq $Italic -and ($Marks[$Marks.Count - 1].Start + $Marks[$Marks.Count - 1].Len) -eq $start) {
                    $Marks[$Marks.Count - 1].Len++
                } else {
                    $Marks.Add([pscustomobject]@{ Start = $start; Len = 1; Bold = $Bold; Italic = $Italic })
                }
            }
            $i++
        }
    }
}

function Get-Marks {
    param([bool]$Bold, [bool]$Italic)
    if ($Bold -and $Italic) { return '***' }
    if ($Bold) { return '**' }
    if ($Italic) { return '*' }
    return ''
}

# 读取段内加粗/斜体格式，还原为 markdown 标记；
# 逐字符读取格式（Range.Runs 在 COM 下不可用），合并相邻同格式为段，
# 天然支持中文段内部分格式；段边界空格放到标记外（**bold** and），段内空格留在标记内（**Shanghai University**）
function Get-MarkedText {
    param([object]$Range)
    $runs = New-Object System.Collections.Generic.List[object]
    foreach ($c in $Range.Characters) {
        $t = $c.Text
        if ($t -eq "`r" -or $t -eq "`n" -or $t -eq [char]7) { continue }
        $runs.Add([pscustomobject]@{ Text = $t; Bold = ($c.Font.Bold -eq -1); Italic = ($c.Font.Italic -eq -1) })
    }
    $merged = New-Object System.Collections.Generic.List[object]
    foreach ($r in $runs) {
        if ($merged.Count -gt 0 -and $merged[$merged.Count - 1].Bold -eq $r.Bold -and $merged[$merged.Count - 1].Italic -eq $r.Italic) {
            $merged[$merged.Count - 1].Text += $r.Text
        } else {
            $merged.Add([pscustomobject]@{ Text = $r.Text; Bold = $r.Bold; Italic = $r.Italic })
        }
    }
    $sb = New-Object System.Text.StringBuilder
    $lastB = $null
    $lastI = $null
    $pending = ''
    foreach ($cur in $merged) {
        $b = $cur.Bold
        $it = $cur.Italic
        $text = $cur.Text
        $lead = ''
        $trail = ''
        if ($text -match '^(\s+)') { $lead = $Matches[1]; $text = $text.Substring($lead.Length) }
        if ($text -match '(\s+)$') { $trail = $Matches[1]; $text = $text.Substring(0, $text.Length - $trail.Length) }
        if ($lead) { $pending += $lead }
        if ($b -ne $lastB -or $it -ne $lastI) {
            if ($null -ne $lastB) { [void]$sb.Append((Get-Marks $lastB $lastI)) }
            if ($pending) { [void]$sb.Append($pending); $pending = '' }
            $marks = Get-Marks $b $it
            if ($marks) {
                $t = $sb.ToString()
                if ($t.Length -gt 0) {
                    $lc = $t[$t.Length - 1]
                    if ($lc -eq '*' -or $lc -match '[A-Za-z0-9]') { [void]$sb.Append(' ') }
                }
                [void]$sb.Append($marks)
            }
            $lastB = $b
            $lastI = $it
        }
        [void]$sb.Append($text)
        if ($trail) { $pending += $trail }
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

        # 不写首段标题（页面和文章一致）——标题由 front-matter title / 模板渲染，
        # 若写进 Word 会让用户删除后重新生成又出现，且删除后 Get-WordTitle 会误取正文首段；
        # （PS 5.1 COM 下 BuiltInDocumentProperties 不可用，不尝试写文档属性）

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
                $r = Add-WordPara -Word $word -Doc $doc -Text $Matches[1] -Style -19   # List Bullet
                $r.ListFormat.ApplyBulletDefault()   # 应用真实项目符号，Word 里可见圆点
                continue
            }
            # 普通段落 + 加粗/斜体（Add-WordPara 解析 ** 与 * 标记）
            Add-WordPara -Word $word -Doc $doc -Text $line -Style -1
        }

        # Documents.Add 的初始空段会留在第一行 → 删除它
        $first = $doc.Paragraphs.Item(1)
        if (-not $first.Range.Text.Trim()) { $first.Range.Delete() }

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
    $fm = ''
    if (Test-Path -LiteralPath $PageMd) {
        $old = Get-Content -LiteralPath $PageMd -Raw -Encoding UTF8
        if ($old -match '(?s)^---\r?\n(.*?)\r?\n---') { $fm = $Matches[1].TrimEnd() }
    }

    $word = New-Object -ComObject Word.Application
    try {
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $doc = $word.Documents.Open($DocxPath, $false, $true)
        $styleH1 = $doc.Styles.Item(-2).NameLocal
        $styleH2 = $doc.Styles.Item(-3).NameLocal
        $styleBullet = $doc.Styles.Item(-19).NameLocal

        $lines = New-Object System.Collections.Generic.List[string]
        $fmTitle = ''
        if ($fm -match '(?m)^title:\s*(.+)$') { $fmTitle = $Matches[1].Trim() }
        # 仅当第一个 H1 文本与 front-matter title 相同（标题重复）时跳过它；
        # 若正文 H1 与标题不同（如 # Zhen Zhang + title: Home）则保留
        $skipTitle = $true
        foreach ($p in $doc.Paragraphs) {
            $name = $p.Style.NameLocal
            if ($name -eq $styleH1 -or $name -eq $styleH2) {
                # 标题整段加粗是样式属性，不还原 ** 标记
                $text = ($p.Range.Text -replace '[\r\n\x07]', '').Trim()
                if (-not $text) { continue }
                if ($name -eq $styleH1) {
                    if ($skipTitle) {
                        $skipTitle = $false
                        if ($text -eq $fmTitle) { continue }
                    }
                    $lines.Add("# $text")
                } else {
                    $lines.Add("## $text")
                }
            } else {
                $text = Get-MarkedText -Range $p.Range
                if (-not $text) { continue }
                # 列表判断：有真实项目符号（ListType>0，用户 Word 里加/去圆点都能识别），
                # 或仍是 List Bullet 样式（旧文档兜底）
                $isList = $false
                try { $isList = ($p.Range.ListFormat.ListType -gt 0) } catch { }
                if (-not $isList -and $name -eq $styleBullet) { $isList = $true }
                if ($isList) { $lines.Add("- $text") } else { $lines.Add($text) }
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
        if (Test-WordEditChanged -DocxPath $d.FullName) { $stale += $base }
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
# 侧边栏姓名优先取站点 _config.yml 的 author，因此同步更新站点与主题两处
function Set-ThemeProfile {
    param([string]$Author, [string]$Bio, [string]$Root)
    $cfg = Join-Path $Root 'themes\Academia\_config.yml'
    $text = [System.IO.File]::ReadAllText($cfg, [System.Text.Encoding]::UTF8)
    $yamlAuthor = "'" + ($Author -replace "'", "''") + "'"
    $yamlBio = "'" + ($Bio -replace "'", "''") + "'"
    $text = $text -replace "(?m)^author:\s*.*$", "author: $yamlAuthor"
    $text = $text -replace "(?m)^author_bio:\s*.*$", "author_bio: $yamlBio"
    [System.IO.File]::WriteAllText($cfg, $text, (New-Object System.Text.UTF8Encoding($false)))
    $site = Join-Path $Root '_config.yml'
    $stext = [System.IO.File]::ReadAllText($site, [System.Text.Encoding]::UTF8)
    if ($stext -match "(?m)^author:") {
        $stext = $stext -replace "(?m)^author:.*$", "author: $yamlAuthor"
    } else {
        $stext = $stext + "`nauthor: $yamlAuthor`n"
    }
    [System.IO.File]::WriteAllText($site, $stext, (New-Object System.Text.UTF8Encoding($false)))
}

# ---------- 读取作者名与简介 ----------
function Get-ThemeProfile {
    param([string]$Root)
    $cfg = Join-Path $Root 'themes\Academia\_config.yml'
    $text = [System.IO.File]::ReadAllText($cfg, [System.Text.Encoding]::UTF8)
    $author = ''
    $bio = ''
    if ($text -match "(?m)^author:\s*'((?:''|[^'])*)'") { $author = $Matches[1] -replace "''", "'" }
    if ($text -match "(?m)^author_bio:\s*'((?:''|[^'])*)'") { $bio = $Matches[1] -replace "''", "'" }
    return @{ Author = $author; Bio = $bio }
}

# ---------- 读取主题导航菜单（themes\Academia\_config.yml 的 menu: 块，按 YAML 顺序） ----------
function Get-MenuOrder {
    param([string]$Root)
    $cfgPath = Join-Path $Root 'themes\Academia\_config.yml'
    $lines = @(Get-Content -LiteralPath $cfgPath -Encoding UTF8)
    $idx = -1
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match '^\s*menu:\s*$') { $idx = $i; break }
    }
    if ($idx -lt 0) { throw '未找到 menu 配置（themes\Academia\_config.yml）' }
    $items = @()
    $j = $idx + 1
    while ($j -lt $lines.Count -and $lines[$j] -match '^(\s{2,})([^:#][^:]*?)\s*:\s*(.*)$') {
        $key = $Matches[2].Trim()
        if ($key -match "^'(.*)'$") { $key = $Matches[1] -replace "''", "'" }
        $items += [pscustomobject]@{ Key = $key; Value = $Matches[3].Trim(); Indent = $Matches[1]; Line = $j }
        $j++
    }
    if ($items.Count -eq 0) { throw 'menu 配置中没有可调整的菜单项' }
    return [pscustomobject]@{ Path = $cfgPath; Lines = $lines; Items = $items }
}

# ---------- 将独立页面加入导航 ----------
function Add-MenuItem {
    param([string]$Root, [string]$Label, [string]$Url)
    $cfg = Get-MenuOrder -Root $Root
    if ($cfg.Items | Where-Object { $_.Key -eq $Label }) { throw "导航项已存在: $Label" }
    $line = "  ${Label}: $Url"
    if ($Label -notmatch '^[A-Za-z0-9 _-]+$') {
        $quoted = "'" + ($Label -replace "'", "''") + "'"
        $line = "  ${quoted}: $Url"
    }
    $insertAt = $cfg.Items[-1].Line + 1
    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($item in $cfg.Lines) { $lines.Add($item) }
    $lines.Insert($insertAt, $line)
    [System.IO.File]::WriteAllText($cfg.Path, ($lines -join "`n"), (New-Object System.Text.UTF8Encoding($false)))
    return $cfg.Path
}

function Remove-MenuItem {
    param([string]$Root, [string]$Label)
    $cfg = Get-MenuOrder -Root $Root
    $item = $cfg.Items | Where-Object { $_.Key -eq $Label } | Select-Object -First 1
    if (-not $item) { return $false }
    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($line in $cfg.Lines) { $lines.Add($line) }
    $lines.RemoveAt($item.Line)
    [System.IO.File]::WriteAllText($cfg.Path, ($lines -join "`n"), (New-Object System.Text.UTF8Encoding($false)))
    return $true
}

# ---------- 按新顺序写回主题导航菜单（可同时改显示标题） ----------
# Order: 新顺序的菜单 key 列表；Rename: 原 key → 新显示标题（可选）
function Set-MenuOrder {
    param([string]$Root, [string[]]$Order, [hashtable]$Rename)
    $cfg = Get-MenuOrder -Root $Root
    if ($Order.Count -ne $cfg.Items.Count) { throw "菜单项数量不匹配（期望 $($cfg.Items.Count)，传入 $($Order.Count)）" }
    $byKey = @{}
    foreach ($it in $cfg.Items) { $byKey[$it.Key] = $it }
    $slots = @($cfg.Items | ForEach-Object { $_.Line })   # 原槽位行号（按原顺序）
    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($ln in $cfg.Lines) { $lines.Add($ln) }
    $newKeys = @()
    for ($n = 0; $n -lt $Order.Count; $n++) {
        $it = $byKey[$Order[$n]]
        if (-not $it) { throw "菜单项不存在: $($Order[$n])" }
        $key = $it.Key
        if ($Rename -and $Rename.ContainsKey($key)) { $key = $Rename[$key] }
        if (-not $key) { throw '菜单显示标题不能为空' }
        $newKeys += $key
        $lines[$slots[$n]] = "$($it.Indent)$($key): $($it.Value)"
    }
    $dups = $newKeys | Group-Object | Where-Object { $_.Count -gt 1 }
    if ($dups) { throw "菜单显示标题重复: $((($dups | ForEach-Object { $_.Name }) -join ', '))" }
    [System.IO.File]::WriteAllText($cfg.Path, ($lines -join "`n"), (New-Object System.Text.UTF8Encoding($false)))
    return $cfg.Path
}

# ---------- Git 发布 ----------
function Publish-Git {
    param([string]$Message, [string]$Root)
    if (-not $Message) { $Message = '更新内容' }
    $previousErrorActionPreference = $ErrorActionPreference
    Push-Location $Root
    try {
        Invoke-HexoBuild -Root $Root | Out-Null
        Test-PublicLinks -Root $Root | Out-Null
        # Git 将换行转换提示写入 stderr；在严格错误模式下需把它当作普通输出，
        # 最终是否成功只根据各命令的退出码判断。
        $ErrorActionPreference = 'Continue'
        $addOutput = @(& git add -A 2>&1)
        if ($LASTEXITCODE -ne 0) { throw "git add 失败:`n$($addOutput -join "`n")" }
        & git diff --cached --quiet
        if ($LASTEXITCODE -eq 0) { return $false }
        $commitOutput = @(& git commit -m $Message 2>&1)
        if ($LASTEXITCODE -ne 0) { throw "git commit 失败:`n$($commitOutput -join "`n")" }
        $branch = (& git branch --show-current).Trim()
        if (-not $branch) { throw '无法确定当前 Git 分支' }
        $pushOutput = @(& git push origin $branch 2>&1)
        if ($LASTEXITCODE -ne 0) { throw "git push 失败:`n$($pushOutput -join "`n")" }
        return $true
    } finally {
        $ErrorActionPreference = $previousErrorActionPreference
        Pop-Location
    }
}





