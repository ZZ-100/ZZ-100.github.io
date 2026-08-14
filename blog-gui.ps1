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
$form.Size = New-Object System.Drawing.Size(850, 600)
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
$lblList.Text = '已有文章（双击编辑）:'
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
    if ($list.Items.Count -gt 0) { $list.Items[0].Selected = $true }
}

# ---------- 文章网页 URL 与本地服务器 ----------
function Get-PostUrl {
    param([string]$FilePath)
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

# --- 底部工具按钮 ---
$btnEdit = New-Object System.Windows.Forms.Button
$btnEdit.Text = '编辑选中'
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

# --- 状态栏 ---
$status = New-Object System.Windows.Forms.Label
$status.Location = New-Object System.Drawing.Point(12, 480)
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
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = 'Word 文档 (*.docx;*.doc)|*.docx;*.doc'
    $ofd.Title = '选择要导入的 Word 文档'
    if ($ofd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
    $status.Text = '正在从 Word 导入，请稍候...'
    try {
        $title = Get-WordTitle -DocxPath $ofd.FileName
        $file = New-PostFromWord -DocxPath $ofd.FileName -Title $title -Root $Root
        $txtTitle.Text = $title
        $status.Text = "导入成功: $title（排版已保留）"
        Refresh-List
        [System.Windows.Forms.MessageBox]::Show("导入成功，文章标题: $title`n排版已保留。", '完成') | Out-Null
        Start-Process notepad $file
    } catch {
        $status.Text = '导入失败'
        [System.Windows.Forms.MessageBox]::Show("导入失败: $($_.Exception.Message)", '错误', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
})

$btnEdit.Add_Click({
    if (-not $list.SelectedItems.Count) { return }
    Start-Process notepad $list.SelectedItems[0].Tag
})

$list.Add_DoubleClick({
    if (-not $list.SelectedItems.Count) { return }
    Start-Process notepad $list.SelectedItems[0].Tag
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

# ---------- 组装 ----------
$form.Controls.AddRange(@($lblTitle, $txtTitle, $btnNew, $btnWordNew, $btnImport, $btnPublish, $lblList, $list, $btnEdit, $btnDelete, $btnPreview, $btnPreviewSel, $btnBuild, $status))
Refresh-List
[System.Windows.Forms.Application]::EnableVisualStyles()
$form.ShowDialog() | Out-Null



