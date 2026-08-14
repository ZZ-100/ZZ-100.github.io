<#
  blog-gui.ps1 - 学术主页图形化管理工具

  功能：
    - 输入标题直接新建文章（自动带 academia: true）
    - 从 Word 文档导入（自动读取文档标题，保留排版）
    - 文章列表管理（双击编辑）
    - 本地预览 / 构建 / 发布

  运行：
    powershell -ExecutionPolicy Bypass -File .\blog-gui.ps1
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $Root 'blog-lib.ps1')

# ---------- 构建表单 ----------
$form = New-Object System.Windows.Forms.Form
$form.Text = '学术主页管理工具'
$form.Size = New-Object System.Drawing.Size(850, 660)
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9)

# --- 顶部：标题输入 + 操作按钮 ---
$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = '文章标题:'
$lblTitle.Location = New-Object System.Drawing.Point(12, 17)
$lblTitle.AutoSize = $true

$txtTitle = New-Object System.Windows.Forms.TextBox
$txtTitle.Location = New-Object System.Drawing.Point(90, 14)
$txtTitle.Size = New-Object System.Drawing.Size(300, 23)

$btnNew = New-Object System.Windows.Forms.Button
$btnNew.Text = '新建文章'
$btnNew.Location = New-Object System.Drawing.Point(400, 12)
$btnNew.Size = New-Object System.Drawing.Size(90, 27)

$btnWordNew = New-Object System.Windows.Forms.Button
$btnWordNew.Text = '用 Word 新建...'
$btnWordNew.Location = New-Object System.Drawing.Point(500, 12)
$btnWordNew.Size = New-Object System.Drawing.Size(110, 27)

$btnImport = New-Object System.Windows.Forms.Button
$btnImport.Text = '从 Word 导入...'
$btnImport.Location = New-Object System.Drawing.Point(620, 12)
$btnImport.Size = New-Object System.Drawing.Size(110, 27)

$btnPublish = New-Object System.Windows.Forms.Button
$btnPublish.Text = '发布'
$btnPublish.Location = New-Object System.Drawing.Point(740, 12)
$btnPublish.Size = New-Object System.Drawing.Size(80, 27)
$btnPublish.BackColor = [System.Drawing.Color]::FromArgb(46, 139, 87)
$btnPublish.ForeColor = [System.Drawing.Color]::White

# --- 文章列表 ---
$lblList = New-Object System.Windows.Forms.Label
$lblList.Text = '文章与页面（双击=Word 编辑，选中后点"预览选中"看网页效果）:'
$lblList.Location = New-Object System.Drawing.Point(12, 50)
$lblList.AutoSize = $true

$list = New-Object System.Windows.Forms.ListView
$list.Location = New-Object System.Drawing.Point(12, 72)
$list.Size = New-Object System.Drawing.Size(810, 360)
$list.View = 'Details'
$list.FullRowSelect = $true
$list.MultiSelect = $false
$list.Columns.Add('标题', 350) | Out-Null
$list.Columns.Add('更新时间', 160) | Out-Null

function Refresh-List {
    $list.Items.Clear()
    $posts = Get-ChildItem -LiteralPath (Join-Path $Root 'source\_posts') -Filter '*.md' -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending
    foreach ($p in $posts) {
        $item = New-Object System.Windows.Forms.ListViewItem($p.BaseName)
        $item.SubItems.Add($p.LastWriteTime.ToString('yyyy-MM-dd HH:mm')) | Out-Null
        $item.Tag = $p.FullName
        $list.Items.Add($item) | Out-Null
    }
    # 独立页面（source 下各目录的 index.md，如 about）
    $pages = Get-ChildItem -LiteralPath (Join-Path $Root 'source') -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -ne '_posts' -and $_.Name -notlike '_*' }
    foreach ($d in $pages) {
        $pageFile = Join-Path $d.FullName 'index.md'
        if (Test-Path -LiteralPath $pageFile) {
            $item = New-Object System.Windows.Forms.ListViewItem($d.Name)
            $item.SubItems.Add((Get-Item -LiteralPath $pageFile).LastWriteTime.ToString('yyyy-MM-dd HH:mm')) | Out-Null
            $item.Tag = $pageFile
            $list.Items.Add($item) | Out-Null
        }
    }
    if ($list.Items.Count -gt 0) { $list.Items[0].Selected = $true }
}

# ---------- 文章网页 URL 与本地服务器 ----------
function Get-PostUrl {
    param([string]$FilePath)
    # 独立页面：source\<dir>\index.md → /<dir>/
    if ($FilePath -match 'source\\([^\\]+)\\index\.md$') {
        return "http://localhost:4000/$($Matches[1])/"
    }
    $raw = Get-Content -LiteralPath $FilePath -Raw -Encoding UTF8
    $date = ''
    if ($raw -match '(?m)^date:\s*(\d{4})-(\d{2})-(\d{2})') {
        $date = "$($Matches[1])/$($Matches[2])/$($Matches[3])"
    }
    $slug = [System.Uri]::EscapeDataString([System.IO.Path]::GetFileNameWithoutExtension($FilePath))
    return "http://localhost:4000/$date/$slug/"
}

function Start-HexoServer {
    $listening = Get-NetTCPConnection -LocalPort 4000 -State Listen -ErrorAction SilentlyContinue
    if (-not $listening) {
        Start-Process powershell -WindowStyle Minimized -ArgumentList '-NoExit', '-Command', "cd $Root; npm.cmd run server"
        for ($i = 0; $i -lt 20; $i++) {
            Start-Sleep -Milliseconds 1000
            $listening = Get-NetTCPConnection -LocalPort 4000 -State Listen -ErrorAction SilentlyContinue
            if ($listening) { break }
        }
    }
}

# 在 Word 中打开文章进行编辑
function Open-PostInWord {
    param([string]$MdPath, [string]$Root)
    $editDir = Join-Path $Root '.word_edits'
    if (-not (Test-Path -LiteralPath $editDir)) { New-Item -ItemType Directory -Path $editDir | Out-Null }
    # 独立页面 source\<dir>\index.md → 用目录名作文件名（about.docx），避免歧义
    if ($MdPath -match 'source\\([^\\]+)\\index\.md$') { $slug = $Matches[1] } else { $slug = [System.IO.Path]::GetFileNameWithoutExtension($MdPath) }
    $docx = Join-Path $editDir ($slug + '.docx')
    # 若 Word 已打开该文档的旧版本，重新生成后打开仍会复用旧窗口，导致误保存旧内容 → 提示先关闭
    if (Get-Process WINWORD -ErrorAction SilentlyContinue) {
        try {
            $w = New-Object -ComObject Word.Application
            $opened = @($w.Documents | Where-Object { $_.FullName -eq $docx })
            if ($opened.Count -gt 0) {
                $w.Quit()
                [System.Windows.Forms.MessageBox]::Show("Word 中已打开 [$slug].docx 的旧版本。`n请先在 Word 中关闭该文档窗口，再重新双击编辑。", '请先关闭 Word 文档', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
                return
            }
            $w.Quit()
        } catch { }
    }
    if (Test-Path -LiteralPath $docx) { Remove-Item -LiteralPath $docx }
    Convert-MdToWord -MdPath $MdPath -DocxPath $docx
    Start-Process $docx
    $status.Text = "已在 Word 打开 [$slug]，改完保存后回到本窗口，选中它再点『从 Word 导入』即可更新"
}

# --- 底部工具按钮（第一行） ---
$btnEdit = New-Object System.Windows.Forms.Button
$btnEdit.Text = 'Word 编辑'
$btnEdit.Location = New-Object System.Drawing.Point(12, 445)
$btnEdit.Size = New-Object System.Drawing.Size(90, 27)

$btnDelete = New-Object System.Windows.Forms.Button
$btnDelete.Text = '删除选中'
$btnDelete.Location = New-Object System.Drawing.Point(110, 445)
$btnDelete.Size = New-Object System.Drawing.Size(90, 27)

$btnPreview = New-Object System.Windows.Forms.Button
$btnPreview.Text = '本地预览'
$btnPreview.Location = New-Object System.Drawing.Point(210, 445)
$btnPreview.Size = New-Object System.Drawing.Size(90, 27)

$btnPreviewSel = New-Object System.Windows.Forms.Button
$btnPreviewSel.Text = '预览选中'
$btnPreviewSel.Location = New-Object System.Drawing.Point(310, 445)
$btnPreviewSel.Size = New-Object System.Drawing.Size(90, 27)

$btnBuild = New-Object System.Windows.Forms.Button
$btnBuild.Text = '构建检查'
$btnBuild.Location = New-Object System.Drawing.Point(410, 445)
$btnBuild.Size = New-Object System.Drawing.Size(90, 27)

# --- 底部工具按钮（第二行：站点设置） ---
$btnAvatar = New-Object System.Windows.Forms.Button
$btnAvatar.Text = '更换头像...'
$btnAvatar.Location = New-Object System.Drawing.Point(12, 480)
$btnAvatar.Size = New-Object System.Drawing.Size(100, 27)

$btnCv = New-Object System.Windows.Forms.Button
$btnCv.Text = '更换CV...'
$btnCv.Location = New-Object System.Drawing.Point(120, 480)
$btnCv.Size = New-Object System.Drawing.Size(100, 27)

$btnSettings = New-Object System.Windows.Forms.Button
$btnSettings.Text = '设置...'
$btnSettings.Location = New-Object System.Drawing.Point(228, 480)
$btnSettings.Size = New-Object System.Drawing.Size(90, 27)

$btnMenu = New-Object System.Windows.Forms.Button
$btnMenu.Text = '导航栏顺序...'
$btnMenu.Location = New-Object System.Drawing.Point(326, 480)
$btnMenu.Size = New-Object System.Drawing.Size(110, 27)

# --- 状态栏 ---
$status = New-Object System.Windows.Forms.Label
$status.Location = New-Object System.Drawing.Point(12, 518)
$status.Size = New-Object System.Drawing.Size(810, 60)
$status.ForeColor = [System.Drawing.Color]::Gray
$status.Text = '就绪'

# ---------- 事件 ----------
$btnNew.Add_Click({
    $title = $txtTitle.Text.Trim()
    if (-not $title) { [System.Windows.Forms.MessageBox]::Show('请先输入文章标题', '提示') | Out-Null; return }
    try {
        $file = New-BlankPost -Title $title -Root $Root
        $status.Text = "已创建: $(Split-Path $file -Leaf)"
        Refresh-List
        Start-Process notepad $file
    } catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, '提示') | Out-Null
    }
})

$btnWordNew.Add_Click({
    $title = $txtTitle.Text.Trim()
    if (-not $title) { $title = '新文章' }
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = 'Word 文档 (*.docx)|*.docx'
    $sfd.FileName = ($title -replace '[\\/:*?"<>|]', '_') + '.docx'
    $sfd.InitialDirectory = [Environment]::GetFolderPath('MyDocuments')
    if ($sfd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
    $status.Text = '正在创建 Word 草稿...'
    try {
        New-WordDraft -Title $title -SavePath $sfd.FileName | Out-Null
        $status.Text = '草稿已创建，在 Word 中编辑后点"从 Word 导入"'
        Start-Process $sfd.FileName
        [System.Windows.Forms.MessageBox]::Show("Word 草稿已创建，样式已自动配置：`n· 标题：黑体小三 居中`n· 一级小标题：黑体四号`n· 正文：宋体小四 1.5倍行距 首行缩进`n`n编辑保存后，点『从 Word 导入』选择此文件即可发布。", '已创建') | Out-Null
    } catch {
        [System.Windows.Forms.MessageBox]::Show("创建失败: $($_.Exception.Message)", '错误', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
})

$btnImport.Add_Click({
    # 优先使用选中项对应的 .word_edits\<名>.docx（双击打开的就是它），免去文件选择
    $docxPath = $null
    if ($list.SelectedItems.Count -gt 0) {
        $selMd = $list.SelectedItems[0].Tag
        if ($selMd -match 'source\\([^\\]+)\\index\.md$') { $slug = $Matches[1] } else { $slug = [System.IO.Path]::GetFileNameWithoutExtension($selMd) }
        $cand = Join-Path $Root ('.word_edits\' + $slug + '.docx')
        if (Test-Path -LiteralPath $cand) { $docxPath = $cand }
    }
    if (-not $docxPath) {
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Filter = 'Word 文档 (*.docx;*.doc)|*.docx;*.doc'
        $ofd.Title = '选择要导入的 Word 文档'
        if ($ofd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $docxPath = $ofd.FileName
    }
    $status.Text = '正在从 Word 导入，请稍候...'
    # 页面往返：docx 文件名匹配 source\<名>\index.md → 直接更新页面
    $pageBase = [System.IO.Path]::GetFileNameWithoutExtension($docxPath)
    $pageMd = Join-Path $Root "source\$pageBase\index.md"
    if (Test-Path -LiteralPath $pageMd) {
        $r = [System.Windows.Forms.MessageBox]::Show("更新页面 [$pageBase] 的内容吗？", '确认更新页面', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($r -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        try {
            Convert-WordToPageMd -DocxPath $docxPath -PageMd $pageMd
            $status.Text = "已更新页面: $pageBase"
            Refresh-List
            [System.Windows.Forms.MessageBox]::Show("页面 [$pageBase] 已更新，点『发布』即可上线。", '完成') | Out-Null
        } catch {
            $status.Text = '页面更新失败'
            [System.Windows.Forms.MessageBox]::Show("页面更新失败: $($_.Exception.Message)", '错误', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        }
        return
    }
try {
        $title = Get-WordTitle -DocxPath $docxPath
        $force = $false
        $safeTitle = $title -replace '[\\/:*?"<>|]', '_'
        # 优先按 docx 文件名映射已有文章（Word 里删除标题段后 Get-WordTitle 会取正文首段）；
        # 已有文章时标题以 md front-matter 为准（Word 文档属性在 PS COM 下不可用）
        $docBase = [System.IO.Path]::GetFileNameWithoutExtension($docxPath)
        $docMd = Join-Path $Root "source\_posts\$docBase.md"
        if (Test-Path -LiteralPath $docMd) {
            $postMd = $docMd
            $c = Get-Content -LiteralPath $docMd -Raw -Encoding UTF8
            if ($c -match '(?m)^title:\s*(.+)$') { $title = $Matches[1].Trim() }
        } else {
            $postMd = Join-Path $Root "source\_posts\$safeTitle.md"
        }
        if (Test-Path -LiteralPath $postMd) {
            $r = [System.Windows.Forms.MessageBox]::Show("文章 [$(Split-Path $postMd -Leaf)] 已存在，是否用此 Word 文档覆盖更新？", '覆盖确认', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($r -ne [System.Windows.Forms.DialogResult]::Yes) { return }
            $force = $true
        }
        $file = New-PostFromWord -DocxPath $docxPath -Title $title -Root $Root -Force:$force -TargetFile $postMd
        $txtTitle.Text = $title
        $status.Text = "导入成功: $title（排版已保留）"
        Refresh-List
        if ($force) {
            [System.Windows.Forms.MessageBox]::Show("已更新文章 [$title]，排版保留。`n点『发布』即可上线。", '更新完成') | Out-Null
        } else {
            [System.Windows.Forms.MessageBox]::Show("导入成功，文章标题: $title`n排版已保留。", '完成') | Out-Null
        }
    } catch {
        $status.Text = '导入失败'
        [System.Windows.Forms.MessageBox]::Show("导入失败: $($_.Exception.Message)", '错误', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
})

$btnEdit.Add_Click({
    if (-not $list.SelectedItems.Count) { return }
    Open-PostInWord -MdPath $list.SelectedItems[0].Tag -Root $Root
})

$list.Add_DoubleClick({
    if (-not $list.SelectedItems.Count) { return }
    Open-PostInWord -MdPath $list.SelectedItems[0].Tag -Root $Root
})

$btnDelete.Add_Click({
    if (-not $list.SelectedItems.Count) { return }
    $name = $list.SelectedItems[0].Text
    $r = [System.Windows.Forms.MessageBox]::Show("确定删除文章 [$name] 吗？", '确认删除', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
    if ($r -eq [System.Windows.Forms.DialogResult]::Yes) {
        Remove-Item -LiteralPath $list.SelectedItems[0].Tag
        $status.Text = "已删除: $name"
        Refresh-List
    }
})

$btnPreview.Add_Click({
    Start-Process powershell -ArgumentList '-NoExit', '-Command', "cd $Root; npm.cmd run server"
    $status.Text = '预览窗口已打开 (http://localhost:4000)'
})

$btnPreviewSel.Add_Click({
    if (-not $list.SelectedItems.Count) {
        [System.Windows.Forms.MessageBox]::Show('请先在列表中选择一篇文章', '提示') | Out-Null
        return
    }
    $url = Get-PostUrl -FilePath $list.SelectedItems[0].Tag
    $status.Text = "正在打开: $url"
    Start-HexoServer
    Start-Process $url
    $status.Text = '已打开浏览器预览该文章'
})

$btnAvatar.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = '图片文件 (*.jpg;*.jpeg;*.png;*.bmp;*.gif)|*.jpg;*.jpeg;*.png;*.bmp;*.gif'
    $ofd.Title = '选择头像图片（建议方形照片，自动裁剪居中并缩放）'
    if ($ofd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
    $status.Text = '正在处理头像...'
    try {
        # 直接写入主题头像源文件，避免与站点 source\img 同名冲突（那会被主题占位图覆盖）
        $imgDir = Join-Path $Root 'themes\Academia\source\img'
        if (-not (Test-Path -LiteralPath $imgDir)) { New-Item -ItemType Directory -Path $imgDir | Out-Null }
        $dest = Join-Path $imgDir 'profile.png'
        $src = [System.Drawing.Image]::FromFile($ofd.FileName)
        try {
            $side = [Math]::Min($src.Width, $src.Height)
            $cropX = [Math]::Floor(($src.Width - $side) / 2)
            $cropY = [Math]::Floor(($src.Height - $side) / 2)
            $size = 400
            $bmp = New-Object System.Drawing.Bitmap($size, $size)
            $g = [System.Drawing.Graphics]::FromImage($bmp)
            try {
                $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
                $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
                $g.Clear([System.Drawing.Color]::White)
                $srcRect = New-Object System.Drawing.Rectangle($cropX, $cropY, $side, $side)
                $dstRect = New-Object System.Drawing.Rectangle(0, 0, $size, $size)
                $g.DrawImage($src, $dstRect, $srcRect, [System.Drawing.GraphicsUnit]::Pixel)
            } finally { $g.Dispose() }
            $bmp.Save($dest, [System.Drawing.Imaging.ImageFormat]::Png)
            $bmp.Dispose()
        } finally { $src.Dispose() }
        $status.Text = '头像已更新，发布后生效'
        [System.Windows.Forms.MessageBox]::Show("头像已更新为 400x400 居中裁剪。`n`n点『发布』后 1-2 分钟线上生效。", '完成') | Out-Null
    } catch {
        $status.Text = '头像处理失败'
        [System.Windows.Forms.MessageBox]::Show("头像处理失败: $($_.Exception.Message)", '错误', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
})

$btnCv.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = 'PDF 文件 (*.pdf)|*.pdf'
    $ofd.Title = '选择 CV 文件（PDF）'
    if ($ofd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
    $status.Text = '正在上传 CV...'
    try {
        Set-ThemeCV -PdfPath $ofd.FileName -Root $Root | Out-Null
        $status.Text = 'CV 已更新，发布后生效'
        [System.Windows.Forms.MessageBox]::Show("CV 已更新，并已启用侧边栏『CV 下载』入口。`n`n点『发布』后 1-2 分钟线上生效。", '完成') | Out-Null
    } catch {
        $status.Text = 'CV 上传失败'
        [System.Windows.Forms.MessageBox]::Show("CV 上传失败: $($_.Exception.Message)", '错误', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
})

$btnSettings.Add_Click({
    $profile = Get-ThemeProfile -Root $Root
    $box = New-Object System.Windows.Forms.Form
    $box.Text = '站点设置'
    $box.Size = New-Object System.Drawing.Size(460, 260)
    $box.StartPosition = 'CenterScreen'
    $box.FormBorderStyle = 'FixedDialog'
    $box.MaximizeBox = $false
    $box.MinimizeBox = $false

    $l1 = New-Object System.Windows.Forms.Label
    $l1.Text = '作者姓名:'
    $l1.Location = New-Object System.Drawing.Point(15, 25)
    $l1.AutoSize = $true

    $tb1 = New-Object System.Windows.Forms.TextBox
    $tb1.Text = $profile.Author
    $tb1.Location = New-Object System.Drawing.Point(95, 21)
    $tb1.Size = New-Object System.Drawing.Size(330, 23)

    $l2 = New-Object System.Windows.Forms.Label
    $l2.Text = '个人简介:'
    $l2.Location = New-Object System.Drawing.Point(15, 60)
    $l2.AutoSize = $true

    $tb2 = New-Object System.Windows.Forms.TextBox
    $tb2.Text = $profile.Bio
    $tb2.Location = New-Object System.Drawing.Point(95, 56)
    $tb2.Size = New-Object System.Drawing.Size(330, 70)
    $tb2.Multiline = $true

    $hint = New-Object System.Windows.Forms.Label
    $hint.Text = '简介会显示在侧边栏头像下方（英文）'
    $hint.Location = New-Object System.Drawing.Point(95, 132)
    $hint.AutoSize = $true
    $hint.ForeColor = [System.Drawing.Color]::Gray

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = '保存'
    $ok.Location = New-Object System.Drawing.Point(95, 170)
    $ok.DialogResult = 'OK'

    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = '取消'
    $cancel.Location = New-Object System.Drawing.Point(190, 170)
    $cancel.DialogResult = 'Cancel'

    $box.Controls.AddRange(@($l1, $tb1, $l2, $tb2, $hint, $ok, $cancel))
    $box.AcceptButton = $ok
    $box.CancelButton = $cancel
    $box.ShowDialog() | Out-Null

    if ($box.DialogResult -ne [System.Windows.Forms.DialogResult]::OK) { return }
    $author = $tb1.Text.Trim()
    $bio = $tb2.Text.Trim()
    if (-not $author) { $author = $profile.Author }
    try {
        $bio = $bio -replace "'", ''
        Set-ThemeProfile -Author $author -Bio $bio -Root $Root | Out-Null
        $status.Text = '设置已保存，发布后生效'
        [System.Windows.Forms.MessageBox]::Show("设置已保存。`n若正在本地预览，请重启预览（关闭后重新点『本地预览』）查看效果。", '完成') | Out-Null
    } catch {
        [System.Windows.Forms.MessageBox]::Show("保存失败: $($_.Exception.Message)", '错误', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
})

$btnBuild.Add_Click({
    $status.Text = '正在构建...'
    & npm.cmd run clean 2>&1 | Out-Null
    $out = & npm.cmd run build 2>&1
    if ($LASTEXITCODE -eq 0) {
        $status.Text = '构建成功'
        [System.Windows.Forms.MessageBox]::Show('构建成功，无错误。', '构建检查') | Out-Null
    } else {
        $status.Text = '构建出错，详见输出'
        [System.Windows.Forms.MessageBox]::Show(($out | Out-String), '构建错误', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
})

$btnPublish.Add_Click({
    $box = New-Object System.Windows.Forms.Form
    $box.Text = '发布确认'
    $box.Size = New-Object System.Drawing.Size(380, 160)
    $box.StartPosition = 'CenterScreen'
    $box.FormBorderStyle = 'FixedDialog'
    $box.MaximizeBox = $false
    $box.MinimizeBox = $false

    $l = New-Object System.Windows.Forms.Label
    $l.Text = '提交说明:'
    $l.Location = New-Object System.Drawing.Point(15, 22)
    $l.AutoSize = $true

    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Text = '更新内容'
    $tb.Location = New-Object System.Drawing.Point(95, 18)
    $tb.Size = New-Object System.Drawing.Size(250, 23)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = '发布'
    $ok.Location = New-Object System.Drawing.Point(95, 70)
    $ok.DialogResult = 'OK'

    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = '取消'
    $cancel.Location = New-Object System.Drawing.Point(190, 70)
    $cancel.DialogResult = 'Cancel'

    $box.Controls.AddRange(@($l, $tb, $ok, $cancel))
    $box.AcceptButton = $ok
    $box.CancelButton = $cancel
    $box.ShowDialog() | Out-Null

if ($box.DialogResult -ne [System.Windows.Forms.DialogResult]::OK) { return }
    $msg = $tb.Text.Trim()

    # 检测 .word_edits 中未导入的 Word 修改，先导入再发布
    $stale = @(Sync-WordEdits -Root $Root)
    if ($stale.Count -gt 0) {
        $names = ($stale | ForEach-Object { "$_.docx" }) -join '、'
        $r = [System.Windows.Forms.MessageBox]::Show("检测到以下 Word 文档有尚未导入的修改：`n$names`n`n是否先导入再发布？", '检测到未导入修改', [System.Windows.Forms.MessageBoxButtons]::YesNoCancel, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($r -eq [System.Windows.Forms.DialogResult]::Cancel) { return }
        if ($r -eq [System.Windows.Forms.DialogResult]::Yes) {
            $status.Text = '正在导入 Word 修改...'
            foreach ($base in $stale) {
                $docx = Join-Path $Root ('.word_edits\' + $base + '.docx')
                $pageMd = Join-Path $Root "source\$base\index.md"
                $postMd = Join-Path $Root "source\_posts\$base.md"
                if (Test-Path -LiteralPath $pageMd) {
                    Convert-WordToPageMd -DocxPath $docx -PageMd $pageMd
                } elseif (Test-Path -LiteralPath $postMd) {
                    $title = Get-WordTitle -DocxPath $docx
                    New-PostFromWord -DocxPath $docx -Title $title -Root $Root -Force -TargetFile $postMd | Out-Null
                }
            }
            Refresh-List
        }
    }

    $status.Text = '正在发布...'
    try {
        Publish-Git -Message $msg -Root $Root | Out-Null
        $status.Text = '发布成功，1-2 分钟后自动部署生效'
        [System.Windows.Forms.MessageBox]::Show('发布成功！GitHub Actions 正在自动部署，1-2 分钟生效。', '发布完成') | Out-Null
    } catch {
        $status.Text = '发布失败'
        [System.Windows.Forms.MessageBox]::Show("发布失败: $_", '错误', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
})

$btnMenu.Add_Click({
    $cfg = $null
    try {
        $cfg = Get-MenuOrder -Root $Root
    } catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, '错误', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        return
    }
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = '调整导航栏顺序'
    $dlg.Size = New-Object System.Drawing.Size(330, 320)
    $dlg.StartPosition = 'CenterParent'
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = '选中一项，用上移/下移调整顺序（确定后发布生效）:'
    $lbl.Location = New-Object System.Drawing.Point(12, 10)
    $lbl.Size = New-Object System.Drawing.Size(290, 30)
    $lst = New-Object System.Windows.Forms.ListBox
    $lst.Location = New-Object System.Drawing.Point(12, 44)
    $lst.Size = New-Object System.Drawing.Size(180, 140)
    foreach ($it in $cfg.Items) { [void]$lst.Items.Add($it.Key) }
    $btnUp = New-Object System.Windows.Forms.Button
    $btnUp.Text = '上移'
    $btnUp.Location = New-Object System.Drawing.Point(210, 60)
    $btnUp.Size = New-Object System.Drawing.Size(90, 27)
    $btnDown = New-Object System.Windows.Forms.Button
    $btnDown.Text = '下移'
    $btnDown.Location = New-Object System.Drawing.Point(210, 96)
    $btnDown.Size = New-Object System.Drawing.Size(90, 27)
    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = '确定'
    $btnOK.Location = New-Object System.Drawing.Point(12, 230)
    $btnOK.Size = New-Object System.Drawing.Size(90, 27)
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = '取消'
    $btnCancel.Location = New-Object System.Drawing.Point(110, 230)
    $btnCancel.Size = New-Object System.Drawing.Size(90, 27)
    $btnUp.Add_Click({
        $i = $lst.SelectedIndex
        if ($i -gt 0) { $t = $lst.Items[$i]; $lst.Items[$i] = $lst.Items[$i - 1]; $lst.Items[$i - 1] = $t; $lst.SelectedIndex = $i - 1 }
    })
    $btnDown.Add_Click({
        $i = $lst.SelectedIndex
        if ($i -ge 0 -and $i -lt $lst.Items.Count - 1) { $t = $lst.Items[$i]; $lst.Items[$i] = $lst.Items[$i + 1]; $lst.Items[$i + 1] = $t; $lst.SelectedIndex = $i + 1 }
    })
    $btnOK.Add_Click({
        $order = @($lst.Items | ForEach-Object { "$_" })
        try {
            Set-MenuOrder -Root $Root -Order $order | Out-Null
            $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $dlg.Close()
        } catch {
            [System.Windows.Forms.MessageBox]::Show("保存失败: $($_.Exception.Message)", '错误', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        }
    })
    $btnCancel.Add_Click({ $dlg.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $dlg.Close() })
    $dlg.Controls.AddRange(@($lbl, $lst, $btnUp, $btnDown, $btnOK, $btnCancel))
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $status.Text = '导航栏顺序已更新，点『发布』生效'
    }
})

# ---------- 组装 ----------
$form.Controls.AddRange(@($lblTitle, $txtTitle, $btnNew, $btnWordNew, $btnImport, $btnPublish, $lblList, $list, $btnEdit, $btnDelete, $btnPreview, $btnPreviewSel, $btnBuild, $btnAvatar, $btnCv, $btnSettings, $btnMenu, $status))
Refresh-List
[System.Windows.Forms.Application]::EnableVisualStyles()
$form.ShowDialog() | Out-Null






