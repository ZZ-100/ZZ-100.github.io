<#
    blog-gui-wpf.ps1 - WPF 版学术主页管理器
    使用现有 blog-lib.ps1 的内容、Word、Hexo 与 Git 功能。
#>

param([switch]$SmokeTest)

$ErrorActionPreference = 'Stop'
$Root = $PSScriptRoot

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml
Add-Type -AssemblyName System.Drawing
. (Join-Path $Root 'blog-lib.ps1')

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="学术主页管理器" Width="1160" Height="760" MinWidth="980" MinHeight="640"
        WindowStartupLocation="CenterScreen" Background="#F5F7FB" FontFamily="Segoe UI, Microsoft YaHei UI" FontSize="14">
  <Window.Resources>
    <SolidColorBrush x:Key="Navy" Color="#17365D" />
    <SolidColorBrush x:Key="Blue" Color="#2563EB" />
    <SolidColorBrush x:Key="Text" Color="#1F2937" />
    <SolidColorBrush x:Key="Muted" Color="#6B7280" />
    <SolidColorBrush x:Key="Card" Color="#FFFFFF" />
    <Style TargetType="Button">
      <Setter Property="Padding" Value="14,8" />
      <Setter Property="Margin" Value="0,0,8,0" />
      <Setter Property="Foreground" Value="#1F2937" />
      <Setter Property="Background" Value="#FFFFFF" />
      <Setter Property="BorderBrush" Value="#D8DEE9" />
      <Setter Property="BorderThickness" Value="1" />
      <Setter Property="Cursor" Value="Hand" />
    </Style>
    <Style x:Key="PrimaryButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
      <Setter Property="Background" Value="#2563EB" />
      <Setter Property="Foreground" Value="White" />
      <Setter Property="BorderBrush" Value="#2563EB" />
    </Style>
    <Style x:Key="SideButton" TargetType="Button">
      <Setter Property="HorizontalContentAlignment" Value="Left" />
      <Setter Property="Padding" Value="18,12" />
      <Setter Property="Margin" Value="0,2,0,2" />
      <Setter Property="Background" Value="Transparent" />
      <Setter Property="BorderThickness" Value="0" />
      <Setter Property="Foreground" Value="#D9E5F7" />
      <Setter Property="FontSize" Value="15" />
    </Style>
    <Style TargetType="TextBox">
      <Setter Property="Padding" Value="8,6" />
      <Setter Property="BorderBrush" Value="#D8DEE9" />
      <Setter Property="BorderThickness" Value="1" />
    </Style>
    <Style TargetType="ComboBox">
      <Setter Property="Padding" Value="6,4" />
      <Setter Property="BorderBrush" Value="#D8DEE9" />
    </Style>
  </Window.Resources>

  <Grid>
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="220" />
      <ColumnDefinition Width="*" />
    </Grid.ColumnDefinitions>

    <Border Grid.Column="0" Background="{StaticResource Navy}">
      <DockPanel LastChildFill="True">
        <StackPanel DockPanel.Dock="Top" Margin="22,26,18,22">
          <TextBlock Text="学术主页管理器" Foreground="White" FontSize="20" FontWeight="SemiBold" />
          <TextBlock Text="Zhen Zhang" Foreground="#AFC4E3" Margin="0,5,0,0" />
        </StackPanel>
        <TextBlock DockPanel.Dock="Bottom" x:Name="SidebarStatus" Text="本地内容管理" Foreground="#AFC4E3" Margin="22,18,18,22" />
        <StackPanel x:Name="Sidebar">
          <Button x:Name="NavDashboard" Content="概览" Tag="Dashboard" Style="{StaticResource SideButton}" />
          <Button x:Name="NavContent" Content="内容管理" Tag="Content" Style="{StaticResource SideButton}" />
          <Button x:Name="NavMedia" Content="媒体与 CV" Tag="Media" Style="{StaticResource SideButton}" />
          <Button x:Name="NavNavigation" Content="导航设置" Tag="Navigation" Style="{StaticResource SideButton}" />
          <Button x:Name="NavSettings" Content="站点设置" Tag="Settings" Style="{StaticResource SideButton}" />
          <Button x:Name="NavPublish" Content="发布中心" Tag="Publish" Style="{StaticResource SideButton}" />
        </StackPanel>
      </DockPanel>
    </Border>

    <Grid Grid.Column="1" Margin="28">
      <Grid.RowDefinitions>
        <RowDefinition Height="*" />
        <RowDefinition Height="Auto" />
      </Grid.RowDefinitions>
      <Grid x:Name="DashboardView">
        <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
        <StackPanel>
          <TextBlock Text="概览" FontSize="30" FontWeight="SemiBold" Foreground="{StaticResource Text}" />
          <TextBlock Text="在这里查看主页状态和待处理事项。" Foreground="{StaticResource Muted}" Margin="0,4,0,20" />
        </StackPanel>
        <UniformGrid Grid.Row="1" Columns="4" Margin="0,0,0,22">
          <Border Background="{StaticResource Card}" CornerRadius="10" Padding="18" Margin="0,0,12,0"><StackPanel><TextBlock Text="内容总数" Foreground="{StaticResource Muted}"/><TextBlock x:Name="StatTotal" Text="0" FontSize="28" FontWeight="SemiBold" Margin="0,8,0,0"/></StackPanel></Border>
          <Border Background="{StaticResource Card}" CornerRadius="10" Padding="18" Margin="0,0,12,0"><StackPanel><TextBlock Text="独立页面" Foreground="{StaticResource Muted}"/><TextBlock x:Name="StatPages" Text="0" FontSize="28" FontWeight="SemiBold" Margin="0,8,0,0"/></StackPanel></Border>
          <Border Background="{StaticResource Card}" CornerRadius="10" Padding="18" Margin="0,0,12,0"><StackPanel><TextBlock Text="Git 变更" Foreground="{StaticResource Muted}"/><TextBlock x:Name="StatChanges" Text="0" FontSize="28" FontWeight="SemiBold" Margin="0,8,0,0"/></StackPanel></Border>
          <Border Background="{StaticResource Card}" CornerRadius="10" Padding="18"><StackPanel><TextBlock Text="构建状态" Foreground="{StaticResource Muted}"/><TextBlock x:Name="StatBuild" Text="未检查" FontSize="20" FontWeight="SemiBold" Margin="0,12,0,0"/></StackPanel></Border>
        </UniformGrid>
        <Grid Grid.Row="2">
          <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="300"/></Grid.ColumnDefinitions>
          <Border Grid.Column="0" Background="{StaticResource Card}" CornerRadius="10" Padding="22" Margin="0,0,16,0">
            <StackPanel>
              <TextBlock Text="快速操作" FontSize="19" FontWeight="SemiBold" />
              <WrapPanel Margin="0,18,0,0">
                <Button x:Name="DashNewArticle" Content="新建文章" Style="{StaticResource PrimaryButton}" />
                <Button x:Name="DashNewPage" Content="新建独立页面" />
                <Button x:Name="DashPreview" Content="本地预览" />
                <Button x:Name="DashBuild" Content="构建检查" />
              </WrapPanel>
              <TextBlock Text="建议在发布前先运行构建检查，工具会自动验证站内链接。" Foreground="{StaticResource Muted}" Margin="0,26,0,0" TextWrapping="Wrap" />
            </StackPanel>
          </Border>
          <Border Grid.Column="1" Background="{StaticResource Card}" CornerRadius="10" Padding="22">
            <StackPanel>
              <TextBlock Text="最近内容" FontSize="19" FontWeight="SemiBold" />
              <ListBox x:Name="RecentList" BorderThickness="0" Margin="0,12,0,0" />
            </StackPanel>
          </Border>
        </Grid>
      </Grid>

      <Grid x:Name="ContentView" Visibility="Collapsed">
        <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
        <StackPanel><TextBlock Text="内容管理" FontSize="30" FontWeight="SemiBold" Foreground="{StaticResource Text}"/><TextBlock Text="搜索、编辑、预览和回收主页内容。" Foreground="{StaticResource Muted}" Margin="0,4,0,18"/></StackPanel>
        <DockPanel Grid.Row="1" Margin="0,0,0,14">
          <StackPanel DockPanel.Dock="Right" Orientation="Horizontal">
            <Button x:Name="ContentNewArticle" Content="新建文章" Style="{StaticResource PrimaryButton}" />
            <Button x:Name="ContentNewPage" Content="新建页面" />
          </StackPanel>
          <TextBox x:Name="SearchBox" Width="260" HorizontalAlignment="Left" ToolTip="搜索标题或路径" />
          <ComboBox x:Name="TypeFilter" Width="130" Margin="10,0,0,0" SelectedIndex="0"><ComboBoxItem Content="全部类型"/><ComboBoxItem Content="主页"/><ComboBoxItem Content="文章"/><ComboBoxItem Content="页面"/></ComboBox>
        </DockPanel>
        <DataGrid x:Name="ContentGrid" Grid.Row="2" AutoGenerateColumns="False" IsReadOnly="True" CanUserAddRows="False" SelectionMode="Single" HeadersVisibility="Column" Background="White" BorderBrush="#D8DEE9">
          <DataGrid.Columns>
            <DataGridTextColumn Header="标题" Binding="{Binding Title}" Width="*" />
            <DataGridTextColumn Header="类型" Binding="{Binding Type}" Width="90" />
            <DataGridTextColumn Header="状态" Binding="{Binding Status}" Width="120" />
            <DataGridTextColumn Header="更新时间" Binding="{Binding Updated}" Width="150" />
          </DataGrid.Columns>
        </DataGrid>
        <StackPanel x:Name="ContentActions" Grid.Row="3" Orientation="Horizontal" Margin="0,12,0,0">
          <Button x:Name="ContentWord" Content="Word 编辑" />
          <Button x:Name="ContentImportWord" Content="导入 Word 修改" />
          <Button x:Name="ContentResetWord" Content="重置 Word" ToolTip="舍弃未导入的 Word 草稿，网页内容保持不变" />
          <Button x:Name="ContentPreview" Content="网页预览" />
          <Button x:Name="ContentRecycle" Content="移入回收站" />
          <TextBlock x:Name="ContentHint" Foreground="{StaticResource Muted}" VerticalAlignment="Center" Margin="10,0,0,0" />
        </StackPanel>
      </Grid>

      <Grid x:Name="MediaView" Visibility="Collapsed">
        <StackPanel>
          <TextBlock Text="媒体与 CV" FontSize="30" FontWeight="SemiBold" Foreground="{StaticResource Text}" />
          <TextBlock Text="管理头像、Favicon 和公开简历文件。" Foreground="{StaticResource Muted}" Margin="0,4,0,24" />
          <Border Background="White" CornerRadius="10" Padding="22" Margin="0,0,0,16">
            <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="160"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
              <Image x:Name="AvatarPreview" Width="130" Height="130" Stretch="UniformToFill" />
              <StackPanel Grid.Column="1" VerticalAlignment="Center"><TextBlock Text="头像" FontSize="20" FontWeight="SemiBold"/><TextBlock Text="保留当前头像，可选择新文件并自动裁剪为方形 PNG。" Foreground="{StaticResource Muted}" Margin="0,8,0,16"/><Button x:Name="ChooseAvatar" Content="更换头像" Width="110" HorizontalAlignment="Left"/></StackPanel>
            </Grid>
          </Border>
          <Border Background="White" CornerRadius="10" Padding="22">
            <StackPanel><TextBlock Text="CV 文件" FontSize="20" FontWeight="SemiBold"/><TextBlock x:Name="CvStatus" Text="当前未启用" Foreground="{StaticResource Muted}" Margin="0,8,0,16"/><Button x:Name="ChooseCv" Content="选择 PDF" Width="110" HorizontalAlignment="Left"/></StackPanel>
          </Border>
        </StackPanel>
      </Grid>

      <Grid x:Name="NavigationView" Visibility="Collapsed">
        <StackPanel>
          <TextBlock Text="导航设置" FontSize="30" FontWeight="SemiBold" Foreground="{StaticResource Text}" />
          <TextBlock Text="拖动能力后续可扩展；当前支持顺序调整和显示名称修改。" Foreground="{StaticResource Muted}" Margin="0,4,0,20" />
          <Border Background="White" CornerRadius="10" Padding="22" Width="560" HorizontalAlignment="Left">
            <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="240"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
              <ListBox x:Name="NavList" Height="220" DisplayMemberPath="Label" />
              <StackPanel Grid.Column="1" Margin="24,0,0,0"><TextBlock Text="显示名称" Foreground="{StaticResource Muted}"/><TextBox x:Name="NavRename" Margin="0,6,0,14"/><WrapPanel><Button x:Name="NavApplyName" Content="应用名称"/><Button x:Name="NavUp" Content="上移"/><Button x:Name="NavDown" Content="下移"/></WrapPanel><Button x:Name="NavSave" Content="保存导航" Style="{StaticResource PrimaryButton}" Width="110" HorizontalAlignment="Left" Margin="0,22,0,0"/></StackPanel>
            </Grid>
          </Border>
        </StackPanel>
      </Grid>

      <Grid x:Name="SettingsView" Visibility="Collapsed">
        <StackPanel>
          <TextBlock Text="站点设置" FontSize="30" FontWeight="SemiBold" Foreground="{StaticResource Text}" />
          <TextBlock Text="这些信息会显示在主页侧栏和页面元数据中。" Foreground="{StaticResource Muted}" Margin="0,4,0,20" />
          <Border Background="White" CornerRadius="10" Padding="22" Width="650" HorizontalAlignment="Left">
            <StackPanel><TextBlock Text="作者姓名" Foreground="{StaticResource Muted}"/><TextBox x:Name="AuthorBox" Margin="0,6,0,16"/><TextBlock Text="个人简介" Foreground="{StaticResource Muted}"/><TextBox x:Name="BioBox" Height="90" TextWrapping="Wrap" AcceptsReturn="True" Margin="0,6,0,16"/><Button x:Name="SaveSettings" Content="保存设置" Style="{StaticResource PrimaryButton}" Width="110" HorizontalAlignment="Left"/></StackPanel>
          </Border>
        </StackPanel>
      </Grid>

      <Grid x:Name="PublishView" Visibility="Collapsed">
        <StackPanel>
          <TextBlock Text="发布中心" FontSize="30" FontWeight="SemiBold" Foreground="{StaticResource Text}" />
          <TextBlock Text="发布前会自动构建并检查站内链接。" Foreground="{StaticResource Muted}" Margin="0,4,0,20" />
          <Border Background="White" CornerRadius="10" Padding="22" Width="700" HorizontalAlignment="Left">
            <StackPanel><TextBlock Text="发布检查" FontSize="20" FontWeight="SemiBold"/><TextBox x:Name="PublishSummary" Text="尚未执行检查" Foreground="{StaticResource Muted}" Margin="0,10,0,16" Padding="0" TextWrapping="Wrap" AcceptsReturn="True" IsReadOnly="True" MinHeight="72" MaxHeight="180" VerticalScrollBarVisibility="Auto" BorderThickness="0" Background="#F8FAFC"/><WrapPanel><Button x:Name="RunBuild" Content="构建与链接检查" Style="{StaticResource PrimaryButton}"/><Button x:Name="RunPublish" Content="检查并发布"/><Button x:Name="OpenOnline" Content="打开线上主页"/><Button x:Name="CopyPublishSummary" Content="复制检查内容"/></WrapPanel><TextBlock Text="发布会推送当前 Git 分支，并由 GitHub Actions 部署。" Foreground="{StaticResource Muted}" Margin="0,22,0,0"/></StackPanel>
          </Border>
        </StackPanel>
      </Grid>

      <Border x:Name="StatusPanel" Grid.Row="1" Background="#E8EDF5" CornerRadius="6" Padding="10" Margin="0,16,0,0"><TextBlock x:Name="StatusBar" Text="就绪" Foreground="#40516B" TextTrimming="CharacterEllipsis" /></Border>
    </Grid>
  </Grid>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

function Get-Control([string]$Name) { return $window.FindName($Name) }
$controls = @{}
foreach ($name in @('DashboardView','ContentView','MediaView','NavigationView','SettingsView','PublishView','StatusPanel','StatusBar','SidebarStatus','StatTotal','StatPages','StatChanges','StatBuild','RecentList','ContentGrid','ContentActions','SearchBox','TypeFilter','ContentHint','NavList','NavRename','AuthorBox','BioBox','PublishSummary','CopyPublishSummary','AvatarPreview','CvStatus','NavDashboard','NavContent','NavMedia','NavNavigation','NavSettings','NavPublish','DashNewArticle','DashNewPage','DashPreview','DashBuild','ContentNewArticle','ContentNewPage','ContentWord','ContentImportWord','ContentResetWord','ContentPreview','ContentRecycle','ChooseAvatar','ChooseCv','NavApplyName','NavUp','NavDown','NavSave','SaveSettings','RunBuild','RunPublish','OpenOnline')) { $controls[$name] = Get-Control $name }

function Set-Status([string]$Message) {
    $controls.StatusBar.Text = $Message
}

function Get-MdTitle([string]$Path) {
    $raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    if ($raw -match '(?m)^title:\s*(.+)$') {
        return $Matches[1].Trim().Trim("'").Trim('"')
    }
    return [System.IO.Path]::GetFileNameWithoutExtension($Path)
}

function Get-PendingChangeCount {
    Push-Location $Root
    try {
        $changes = @(& git status --porcelain 2>$null)
        if ($LASTEXITCODE -ne 0) { return '?' }
        return $changes.Count
    } finally {
        Pop-Location
    }
}

function Get-WordEditPath([string]$MdPath) {
    if ($MdPath -match 'source\\([^\\]+)\\index\.md$') {
        $slug = $Matches[1]
    } else {
        $slug = [System.IO.Path]::GetFileNameWithoutExtension($MdPath)
    }
    return Join-Path $Root ('.word_edits\' + $slug + '.docx')
}

function Get-ContentStatus([string]$MdPath) {
    $wordPath = Get-WordEditPath -MdPath $MdPath
    if ((Test-Path -LiteralPath $wordPath) -and (Test-WordEditChanged -DocxPath $wordPath)) {
        return 'Word 待导入'
    }
    return '已保存'
}

function Get-ContentRecords {
    $records = New-Object System.Collections.Generic.List[object]
    $postDir = Join-Path $Root 'source\_posts'
    foreach ($file in (Get-ChildItem -LiteralPath $postDir -Filter '*.md' -File -ErrorAction SilentlyContinue)) {
        $title = Get-MdTitle $file.FullName
        $type = if ($file.BaseName -eq 'home') { '主页' } else { '文章' }
        $updated = $file.LastWriteTime
        $record = [pscustomobject]@{ Title=$title; Type=$type; Status=(Get-ContentStatus -MdPath $file.FullName); Updated=$updated; Path=$file.FullName }
        $records.Add($record)
    }
    foreach ($dir in (Get-ChildItem -LiteralPath (Join-Path $Root 'source') -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -notlike '_*' })) {
        $file = Join-Path $dir.FullName 'index.md'
        if (Test-Path -LiteralPath $file) {
            $updated = (Get-Item -LiteralPath $file).LastWriteTime
            $record = [pscustomobject]@{ Title=(Get-MdTitle $file); Type='页面'; Status=(Get-ContentStatus -MdPath $file); Updated=$updated; Path=$file }
            $records.Add($record)
        }
    }
    return @($records | Sort-Object Updated -Descending)
}

$script:AllRecords = @()
function Apply-ContentFilter {
    $query = $controls.SearchBox.Text.Trim().ToLowerInvariant()
    $type = if ($controls.TypeFilter.SelectedItem) { $controls.TypeFilter.SelectedItem.Content } else { '全部类型' }
    $filtered = @($script:AllRecords | Where-Object {
        (($type -eq '全部类型') -or $_.Type -eq $type) -and (($query -eq '') -or $_.Title.ToLowerInvariant().Contains($query) -or $_.Path.ToLowerInvariant().Contains($query))
    })
    $controls.ContentGrid.ItemsSource = $filtered
    $controls.ContentHint.Text = "$($filtered.Count) 项"
}

function Refresh-Content {
    $script:AllRecords = Get-ContentRecords
    Apply-ContentFilter
    $controls.StatTotal.Text = "$($script:AllRecords.Count)"
    $controls.StatPages.Text = "$(@($script:AllRecords | Where-Object Type -eq '页面').Count)"
    $controls.StatChanges.Text = "$(Get-PendingChangeCount)"
    $recent = @($script:AllRecords | Select-Object -First 5 | ForEach-Object { "$($_.Type)  $($_.Title)" })
    $controls.RecentList.ItemsSource = $recent
}

function Show-View([string]$Name) {
    foreach ($key in @('Dashboard','Content','Media','Navigation','Settings','Publish')) { $controls["${key}View"].Visibility = if ($key -eq $Name) { 'Visible' } else { 'Collapsed' } }
    $labels = @{ Dashboard='概览'; Content='内容管理'; Media='媒体与 CV'; Navigation='导航设置'; Settings='站点设置'; Publish='发布中心' }
    Set-Status("当前页面：$($labels[$Name])")
}

function Read-TextInput([string]$Title, [string]$Prompt, [string]$Initial='') {
    $dialog = New-Object System.Windows.Window
    $dialog.Title = $Title; $dialog.Owner = $window; $dialog.WindowStartupLocation = 'CenterOwner'; $dialog.Width = 420; $dialog.Height = 180; $dialog.ResizeMode = 'NoResize'; $dialog.Background = [Windows.Media.Brushes]::White
    $panel = New-Object System.Windows.Controls.StackPanel; $panel.Margin = New-Object Windows.Thickness(20)
    $label = New-Object System.Windows.Controls.TextBlock; $label.Text = $Prompt; $label.Margin = New-Object Windows.Thickness(0,0,0,8)
    $box = New-Object System.Windows.Controls.TextBox; $box.Text = $Initial; $box.Margin = New-Object Windows.Thickness(0,0,0,16)
    $buttons = New-Object System.Windows.Controls.StackPanel; $buttons.Orientation = 'Horizontal'; $buttons.HorizontalAlignment = 'Right'
    $ok = New-Object System.Windows.Controls.Button; $ok.Content = '确定'; $ok.Width = 80; $ok.IsDefault = $true
    $cancel = New-Object System.Windows.Controls.Button; $cancel.Content = '取消'; $cancel.Width = 80
    $ok.Add_Click({ $dialog.DialogResult = $true }); $cancel.Add_Click({ $dialog.DialogResult = $false })
    $buttons.Children.Add($ok) | Out-Null; $buttons.Children.Add($cancel) | Out-Null; $panel.Children.Add($label) | Out-Null; $panel.Children.Add($box) | Out-Null; $panel.Children.Add($buttons) | Out-Null; $dialog.Content = $panel
    if ($dialog.ShowDialog()) { return $box.Text.Trim() }
    return $null
}

function New-ArticleFromUi {
    $title = Read-TextInput '新建文章' '文章标题'
    if (-not $title) { return }
    try { $file = New-BlankPost -Title $title -Root $Root; Refresh-Content; Show-View 'Content'; Set-Status("文章已创建：$title"); Start-Process notepad $file } catch { [Windows.MessageBox]::Show($_.Exception.Message, '创建失败') | Out-Null }
}

function New-PageFromUi {
    $title = Read-TextInput '新建独立页面' '页面标题'
    if (-not $title) { return }
    try {
        $file = New-BlankPage -Title $title -Root $Root
        $slug = Split-Path -Leaf (Split-Path -Parent $file)
        $answer = [Windows.MessageBox]::Show("页面已创建。是否加入顶部导航？", '更新导航', [Windows.MessageBoxButton]::YesNo, [Windows.MessageBoxImage]::Question)
        if ($answer -eq [Windows.MessageBoxResult]::Yes) { Add-MenuItem -Root $Root -Label $title -Url "/$slug/" | Out-Null }
        Refresh-Content; Show-View 'Content'; Set-Status("页面已创建：$title"); Start-Process notepad $file
    } catch { [Windows.MessageBox]::Show($_.Exception.Message, '创建失败') | Out-Null }
}

function Get-LocalContentUrl([string]$FilePath) {
    if ($FilePath -match 'source\\([^\\]+)\\index\.md$') { return "http://localhost:4000/$($Matches[1])/" }
    $raw = Get-Content -LiteralPath $FilePath -Raw -Encoding UTF8
    if ($raw -match '(?m)^permalink:\s*(\S+)') { $permalink = $Matches[1].Trim().Trim("'").Trim('"'); if ($permalink.StartsWith('/')) { return "http://localhost:4000$permalink" } }
    $date = ''; if ($raw -match '(?m)^date:\s*(\d{4})-(\d{2})-(\d{2})') { $date = "$($Matches[1])/$($Matches[2])/$($Matches[3])" }
    return "http://localhost:4000/$date/$([System.Uri]::EscapeDataString([System.IO.Path]::GetFileNameWithoutExtension($FilePath)))/"
}

function Test-PreviewServer {
    $client = New-Object System.Net.Sockets.TcpClient
    try {
        $task = $client.ConnectAsync('127.0.0.1', 4000)
        return $task.Wait(350) -and $client.Connected
    } catch {
        return $false
    } finally {
        $client.Close()
    }
}

function Start-LocalPreview {
    $listening = Test-PreviewServer
    if (-not $listening) {
        Start-Process powershell.exe -WindowStyle Minimized -WorkingDirectory $Root -ArgumentList '-NoExit','-Command','npm.cmd run server'
        for ($i=0; $i -lt 24; $i++) {
            Start-Sleep -Milliseconds 500
            $listening = Test-PreviewServer
            if ($listening) { break }
        }
    }
    if (-not $listening) { throw '本地预览服务器启动超时，请查看预览窗口中的错误信息。' }
    $controls.SidebarStatus.Text = '本地预览运行中 · 4000'
}

function Open-HomePreview {
    try {
        Start-LocalPreview
        Start-Process 'http://localhost:4000/'
        Set-Status '已打开主页本地预览'
    } catch {
        [Windows.MessageBox]::Show($_.Exception.Message, '预览失败') | Out-Null
    }
}

function Open-SelectedPreview {
    $item = $controls.ContentGrid.SelectedItem
    if (-not $item) { [Windows.MessageBox]::Show('请先选择一项内容。', '提示') | Out-Null; return }
    try { Start-LocalPreview; Start-Process (Get-LocalContentUrl $item.Path); Set-Status '已打开本地网页预览' } catch { [Windows.MessageBox]::Show($_.Exception.Message, '预览失败') | Out-Null }
}

function Open-SelectedWord {
    $item = $controls.ContentGrid.SelectedItem
    if (-not $item) { [Windows.MessageBox]::Show('请先选择一项内容。', '提示') | Out-Null; return }
    $editDir = Join-Path $Root '.word_edits'; New-Item -ItemType Directory -Path $editDir -Force | Out-Null
    $docx = Get-WordEditPath -MdPath $item.Path
    $slug = [System.IO.Path]::GetFileNameWithoutExtension($docx)
    try {
        if (Test-Path -LiteralPath $docx) {
            if (Test-WordEditChanged -DocxPath $docx) { [Windows.MessageBox]::Show('Word 文档有未导入修改，请先导入，再重新打开。', '请先同步') | Out-Null; return }
            Remove-Item -LiteralPath $docx -Force
            $baseline = Get-WordEditBaselinePath $docx; if (Test-Path -LiteralPath $baseline) { Remove-Item -LiteralPath $baseline -Force }
        }
        Convert-MdToWord -MdPath $item.Path -DocxPath $docx | Out-Null; Set-WordEditBaseline $docx; Start-Process $docx; Set-Status "已在 Word 打开：$slug"
    } catch { [Windows.MessageBox]::Show($_.Exception.Message, 'Word 编辑失败') | Out-Null }
}

function Import-SelectedWord {
    $item = $controls.ContentGrid.SelectedItem
    if (-not $item) { [Windows.MessageBox]::Show('请先选择一项内容。', '提示') | Out-Null; return }
    $docx = Get-WordEditPath -MdPath $item.Path
    if (-not (Test-Path -LiteralPath $docx)) {
        [Windows.MessageBox]::Show('没有找到对应的 Word 编辑文件。请先点击“Word 编辑”。', '没有可导入文件') | Out-Null
        return
    }
    if (-not (Test-WordEditChanged -DocxPath $docx)) {
        [Windows.MessageBox]::Show('Word 文件没有检测到新的已保存修改。', '无需导入') | Out-Null
        return
    }
    $answer = [Windows.MessageBox]::Show('请确认已在 Word 中保存修改。现在导入并更新网页内容吗？', '导入 Word 修改', [Windows.MessageBoxButton]::YesNo, [Windows.MessageBoxImage]::Question)
    if ($answer -ne [Windows.MessageBoxResult]::Yes) { return }
    try {
        Set-Status '正在导入 Word 修改…'
        if ($item.Type -eq '页面') {
            Convert-WordToPageMd -DocxPath $docx -PageMd $item.Path | Out-Null
        } else {
            New-PostFromWord -DocxPath $docx -Title $item.Title -Root $Root -Force -TargetFile $item.Path | Out-Null
        }
        Set-WordEditBaseline -DocxPath $docx
        Refresh-Content
        Set-Status "Word 修改已导入：$($item.Title)"
        [Windows.MessageBox]::Show("[$($item.Title)] 已更新。", '导入完成') | Out-Null
    } catch {
        Set-Status 'Word 修改导入失败'
        [Windows.MessageBox]::Show($_.Exception.Message, '导入失败') | Out-Null
    }
}

function Reset-SelectedWord {
    $item = $controls.ContentGrid.SelectedItem
    if (-not $item) { [Windows.MessageBox]::Show('请先选择一项内容。', '提示') | Out-Null; return }
    $docx = Get-WordEditPath -MdPath $item.Path
    if (-not (Test-Path -LiteralPath $docx)) {
        [Windows.MessageBox]::Show('该内容没有 Word 草稿。', '无需重置') | Out-Null
        return
    }
    $answer = [Windows.MessageBox]::Show('确定舍弃该 Word 草稿及未导入修改吗？网页内容不会改变，草稿会移入 .trash，可手动恢复。', '重置 Word', [Windows.MessageBoxButton]::YesNo, [Windows.MessageBoxImage]::Warning)
    if ($answer -ne [Windows.MessageBoxResult]::Yes) { return }
    try {
        $moved = New-Object System.Collections.Generic.List[string]
        $moved.Add((Move-ToRecycleBin -FilePath $docx -Root $Root))
        $baseline = Get-WordEditBaselinePath -DocxPath $docx
        if (Test-Path -LiteralPath $baseline) { $moved.Add((Move-ToRecycleBin -FilePath $baseline -Root $Root)) }
        Refresh-Content
        Set-Status "Word 草稿已重置：$($item.Title)"
        [Windows.MessageBox]::Show("Word 草稿已移入回收目录。`n$($moved -join "`n")", '重置完成') | Out-Null
    } catch {
        [Windows.MessageBox]::Show($_.Exception.Message, '重置失败') | Out-Null
    }
}

function Recycle-Selected {
    $item = $controls.ContentGrid.SelectedItem
    if (-not $item) { [Windows.MessageBox]::Show('请先选择一项内容。', '提示') | Out-Null; return }
    $answer = [Windows.MessageBox]::Show("确定将 [$($item.Title)] 移入回收目录吗？", '确认回收', [Windows.MessageBoxButton]::YesNo, [Windows.MessageBoxImage]::Warning)
    if ($answer -ne [Windows.MessageBoxResult]::Yes) { return }
    try {
        $menuItem = $null
        $targetPath = $item.Path
        if ($item.Path -match 'source\\([^\\]+)\\index\.md$') {
            $pageSlug = $Matches[1]
            $menu = Get-MenuOrder -Root $Root
            $menuItem = $menu.Items | Where-Object { $_.Value.TrimEnd('/') -eq "/$pageSlug" } | Select-Object -First 1
            if ($menuItem) {
                $sync = [Windows.MessageBox]::Show('该页面仍在顶部导航中。继续回收会同时移除导航入口，是否继续？', '同步导航', [Windows.MessageBoxButton]::YesNo, [Windows.MessageBoxImage]::Question)
                if ($sync -ne [Windows.MessageBoxResult]::Yes) { return }
                Remove-MenuItem -Root $Root -Label $menuItem.Key | Out-Null
            }
            $targetPath = Split-Path -Parent $item.Path
        }

        $moved = New-Object System.Collections.Generic.List[string]
        $moved.Add((Move-ToRecycleBin -FilePath $targetPath -Root $Root))
        $wordPath = Get-WordEditPath -MdPath $item.Path
        foreach ($extra in @($wordPath, (Get-WordEditBaselinePath -DocxPath $wordPath))) {
            if (Test-Path -LiteralPath $extra) { $moved.Add((Move-ToRecycleBin -FilePath $extra -Root $Root)) }
        }
        Refresh-Content
        Load-Navigation
        Set-Status "已回收：$($item.Title)"
        [Windows.MessageBox]::Show("内容已移入项目的 .trash 回收目录。`n$($moved -join "`n")", '已回收') | Out-Null
    } catch {
        if ($menuItem -and (Test-Path -LiteralPath $item.Path)) {
            try { Add-MenuItem -Root $Root -Label $menuItem.Key -Url $menuItem.Value | Out-Null } catch { }
        }
        [Windows.MessageBox]::Show($_.Exception.Message, '回收失败') | Out-Null
    }
}

function Load-Media {
    $avatar = Join-Path $Root 'themes\Academia\source\img\profile.png'
    if (Test-Path -LiteralPath $avatar) {
        $bitmap = New-Object Windows.Media.Imaging.BitmapImage
        $bitmap.BeginInit()
        $bitmap.CacheOption = [Windows.Media.Imaging.BitmapCacheOption]::OnLoad
        $bitmap.CreateOptions = [Windows.Media.Imaging.BitmapCreateOptions]::IgnoreImageCache
        $bitmap.UriSource = [Uri]$avatar
        $bitmap.EndInit()
        $bitmap.Freeze()
        $controls.AvatarPreview.Source = $bitmap
    }
    $cv = Join-Path $Root 'source\attaches\CV.pdf'
    $controls.CvStatus.Text = if (Test-Path -LiteralPath $cv) { 'CV 已存在：' + $cv } else { '当前未启用 CV' }
}

function Choose-AvatarFile {
    $dialog = New-Object Microsoft.Win32.OpenFileDialog; $dialog.Filter = '图片文件|*.jpg;*.jpeg;*.png;*.bmp;*.gif'; $dialog.Title = '选择头像'
    if (-not $dialog.ShowDialog()) { return }
    try {
        $imgDir = Join-Path $Root 'themes\Academia\source\img'; New-Item -ItemType Directory -Path $imgDir -Force | Out-Null; $dest = Join-Path $imgDir 'profile.png'; $src = [Drawing.Image]::FromFile($dialog.FileName)
        try { $side=[Math]::Min($src.Width,$src.Height); $x=[Math]::Floor(($src.Width-$side)/2); $y=[Math]::Floor(($src.Height-$side)/2); $bmp=New-Object Drawing.Bitmap(400,400); $g=[Drawing.Graphics]::FromImage($bmp); try { $g.InterpolationMode='HighQualityBicubic'; $g.DrawImage($src,(New-Object Drawing.Rectangle(0,0,400,400)),(New-Object Drawing.Rectangle($x,$y,$side,$side)),[Drawing.GraphicsUnit]::Pixel) } finally { $g.Dispose() }; $bmp.Save($dest,[Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose() } finally { $src.Dispose() }
        Load-Media; Set-Status '头像已更新，发布后生效'
    } catch { [Windows.MessageBox]::Show($_.Exception.Message, '头像处理失败') | Out-Null }
}

function Choose-CvFile {
    $dialog = New-Object Microsoft.Win32.OpenFileDialog; $dialog.Filter = 'PDF 文件|*.pdf'; $dialog.Title = '选择 CV'
    if (-not $dialog.ShowDialog()) { return }
    try { Set-ThemeCV -PdfPath $dialog.FileName -Root $Root | Out-Null; Load-Media; Set-Status 'CV 已更新，发布后生效' } catch { [Windows.MessageBox]::Show($_.Exception.Message, 'CV 更新失败') | Out-Null }
}

$script:NavRows = @()
function Load-Navigation {
    $cfg = Get-MenuOrder -Root $Root
    $script:NavRows = @($cfg.Items | ForEach-Object { [pscustomobject]@{ Key=$_.Key; Label=$_.Key } })
    $controls.NavList.ItemsSource = $null; $controls.NavList.ItemsSource = $script:NavRows
}

function Load-Settings {
    $profile = Get-ThemeProfile -Root $Root; $controls.AuthorBox.Text=$profile.Author; $controls.BioBox.Text=$profile.Bio
}

$script:IsBusy = $false
$script:ActiveProcess = $null
$script:ActiveTimer = $null
function Set-BackgroundBusy([bool]$Busy) {
    $script:IsBusy = $Busy
    foreach ($button in @($controls.DashBuild, $controls.RunBuild, $controls.RunPublish)) {
        $button.IsEnabled = -not $Busy
    }
}

function ConvertFrom-Base64Text([string]$Value) {
    if (-not $Value) { return '' }
    try { return [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($Value)) }
    catch { return '' }
}

function Invoke-WorkerProcess([string]$Action, [string]$Message, [scriptblock]$Completed) {
    Set-BackgroundBusy $true
    $resultPath = [IO.Path]::GetTempFileName()
    $resultPathBase64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($resultPath))
    $workerPath = Join-Path $Root 'blog-worker.ps1'
    $powerShellPath = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $arguments = '-NoLogo -NoProfile -ExecutionPolicy Bypass -File "' + $workerPath + '" -Action ' + $Action + ' -ResultPathBase64 ' + $resultPathBase64
    if ($Message) {
        $messageBase64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($Message))
        $arguments += ' -MessageBase64 ' + $messageBase64
    }

    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = $powerShellPath
    $startInfo.Arguments = $arguments
    $startInfo.WorkingDirectory = $Root
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $startInfo
    $doneBlock = $Completed.GetNewClosure()

    try {
        if (-not $process.Start()) { throw '无法启动后台任务进程' }
    } catch {
        [IO.File]::Delete($resultPath)
        Set-BackgroundBusy $false
        throw
    }

    $timer = New-Object Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromMilliseconds(250)
    $timer.Add_Tick(({
        if (-not $process.HasExited) { return }
        $timer.Stop()
        try {
            $json = if (Test-Path -LiteralPath $resultPath) { [IO.File]::ReadAllText($resultPath, [Text.Encoding]::UTF8).Trim() } else { '' }
            if (-not $json) { throw "后台任务未返回结果（退出码 $($process.ExitCode)）" }
            $payload = $json | ConvertFrom-Json
            $result = [pscustomobject]@{
                Ok = [bool]$payload.Success
                Value = [bool]$payload.Changed
                Message = ConvertFrom-Base64Text $payload.MessageBase64
                Error = ConvertFrom-Base64Text $payload.MessageBase64
                Detail = ConvertFrom-Base64Text $payload.DetailBase64
                ExitCode = $process.ExitCode
            }
        } catch {
            $result = [pscustomobject]@{ Ok=$false; Value=$false; Error=$_.Exception.Message; Detail=($_ | Out-String); ExitCode=$process.ExitCode }
        } finally {
            [IO.File]::Delete($resultPath)
            $process.Dispose()
        }
        try { & $doneBlock $result }
        finally {
            Set-BackgroundBusy $false
            $script:ActiveProcess = $null
            $script:ActiveTimer = $null
        }
    }).GetNewClosure())
    $script:ActiveProcess = $process
    $script:ActiveTimer = $timer
    $timer.Start()
}

function Run-BuildCheck {
    if ($script:IsBusy) { [Windows.MessageBox]::Show('已有构建或发布任务正在运行。', '请稍候') | Out-Null; return }
    $controls.StatBuild.Text='检查中'; $controls.PublishSummary.Text='正在清理、构建并检查内部链接…'; Set-Status '正在执行构建检查…'
    try {
        Invoke-WorkerProcess -Action Build -Completed { param($result); if ($result.Ok) { $controls.StatBuild.Text='通过'; $controls.PublishSummary.Text='构建与链接检查通过'; Set-Status '构建检查通过' } else { $controls.StatBuild.Text='失败'; $controls.PublishSummary.Text=($result.Error + $(if($result.Detail){"`n`n" + $result.Detail}else{''})); Set-Status '构建检查失败' } }
    } catch {
        $controls.StatBuild.Text='失败'; $controls.PublishSummary.Text=$_.Exception.Message; Set-Status '构建检查启动失败'
    }
}

function Run-Publish {
    if ($script:IsBusy) { [Windows.MessageBox]::Show('已有构建或发布任务正在运行。', '请稍候') | Out-Null; return }
    $message = Read-TextInput '发布主页' '提交说明' '更新内容'; if (-not $message) { return }
    $stale = @(Sync-WordEdits -Root $Root); if ($stale.Count -gt 0) { [Windows.MessageBox]::Show("检测到未导入 Word 修改：$($stale -join '、')。请先同步后再发布。", '无法发布') | Out-Null; return }
    $controls.PublishSummary.Text='正在构建并发布…'; Set-Status '正在发布…'
    try {
        Invoke-WorkerProcess -Action Publish -Message $message -Completed { param($result); if ($result.Ok -and $result.Value) { $controls.StatBuild.Text='已发布'; $controls.PublishSummary.Text='发布成功，GitHub Actions 正在部署。'; Set-Status '发布成功' } elseif ($result.Ok) { $controls.PublishSummary.Text='没有需要发布的改动。'; Set-Status '没有待发布改动' } else { $controls.PublishSummary.Text=($result.Error + $(if($result.Detail){"`n`n" + $result.Detail}else{''})); Set-Status '发布失败' } }
    } catch {
        $controls.PublishSummary.Text=$_.Exception.Message; Set-Status '发布任务启动失败'
    }
}

function Run-UiSmokeTest {
    $missing = @($controls.GetEnumerator() | Where-Object { $null -eq $_.Value } | ForEach-Object { $_.Key })
    if ($missing.Count -gt 0) { throw "界面控件未加载：$($missing -join '、')" }

    $window.ShowInTaskbar = $false
    $window.Opacity = 0
    $window.Show()
    try {
        Show-View 'Content'
        $origin = New-Object Windows.Point(0, 0)
        $sizes = @([pscustomobject]@{ Width=1160; Height=760 }, [pscustomobject]@{ Width=980; Height=640 })
        foreach ($size in $sizes) {
            $window.Width = $size.Width
            $window.Height = $size.Height
            $window.UpdateLayout()
            $gridTop = $controls.ContentGrid.TranslatePoint($origin, $window).Y
            $gridBottom = $gridTop + $controls.ContentGrid.ActualHeight
            $actionsTop = $controls.ContentActions.TranslatePoint($origin, $window).Y
            $actionsBottom = $actionsTop + $controls.ContentActions.ActualHeight
            $statusTop = $controls.StatusPanel.TranslatePoint($origin, $window).Y
            if ($gridBottom -gt ($actionsTop + 0.5)) { throw "内容表格与操作按钮发生重叠（$($size.Width)x$($size.Height)）" }
            if ($actionsBottom -gt ($statusTop + 0.5)) { throw "内容操作按钮与状态栏发生重叠（$($size.Width)x$($size.Height)）" }
            if ($controls.StatusPanel.ActualHeight -lt 1) { throw '状态栏没有正确布局' }
        }
        $avatar = Join-Path $Root 'themes\Academia\source\img\profile.png'
        if (Test-Path -LiteralPath $avatar) {
            $probe = [IO.File]::Open($avatar, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::None)
            $probe.Dispose()
        }
        Write-Output "WPF smoke test passed: controls=$($controls.Count), sizes=1160x760/980x640, layout=no-overlap, avatar=unlocked"
    } finally {
        $window.Close()
    }
}

foreach ($button in @($controls.NavDashboard,$controls.NavContent,$controls.NavMedia,$controls.NavNavigation,$controls.NavSettings,$controls.NavPublish)) { $button.Add_Click({ param($sender,$event); Show-View $sender.Tag }) }
$controls.DashNewArticle.Add_Click({ New-ArticleFromUi }); $controls.ContentNewArticle.Add_Click({ New-ArticleFromUi })
$controls.DashNewPage.Add_Click({ New-PageFromUi }); $controls.ContentNewPage.Add_Click({ New-PageFromUi })
$controls.DashPreview.Add_Click({ Open-HomePreview }); $controls.ContentPreview.Add_Click({ Open-SelectedPreview })
$controls.ContentWord.Add_Click({ Open-SelectedWord }); $controls.ContentImportWord.Add_Click({ Import-SelectedWord }); $controls.ContentResetWord.Add_Click({ Reset-SelectedWord }); $controls.ContentRecycle.Add_Click({ Recycle-Selected })
$controls.DashBuild.Add_Click({ Show-View 'Publish'; Run-BuildCheck }); $controls.RunBuild.Add_Click({ Run-BuildCheck }); $controls.RunPublish.Add_Click({ Run-Publish })
$controls.CopyPublishSummary.Add_Click({ try { [Windows.Clipboard]::SetText($controls.PublishSummary.Text); Set-Status '发布检查内容已复制' } catch { [Windows.MessageBox]::Show($_.Exception.Message, '复制失败') | Out-Null } })
$controls.OpenOnline.Add_Click({ Start-Process 'https://zz-100.github.io/' }); $controls.SearchBox.Add_TextChanged({ Apply-ContentFilter }); $controls.TypeFilter.Add_SelectionChanged({ Apply-ContentFilter })
$controls.ContentGrid.Add_MouseDoubleClick({ Open-SelectedWord })
$controls.ChooseAvatar.Add_Click({ Choose-AvatarFile }); $controls.ChooseCv.Add_Click({ Choose-CvFile })
$controls.SaveSettings.Add_Click({ try { Set-ThemeProfile -Author $controls.AuthorBox.Text.Trim() -Bio $controls.BioBox.Text.Trim() -Root $Root | Out-Null; Set-Status '站点设置已保存，发布后生效' } catch { [Windows.MessageBox]::Show($_.Exception.Message, '保存失败') | Out-Null } })
$controls.NavList.Add_SelectionChanged({ if ($controls.NavList.SelectedItem) { $controls.NavRename.Text=$controls.NavList.SelectedItem.Label } })
$controls.NavApplyName.Add_Click({ if ($controls.NavList.SelectedItem -and $controls.NavRename.Text.Trim()) { $controls.NavList.SelectedItem.Label=$controls.NavRename.Text.Trim(); $controls.NavList.Items.Refresh() } })
$controls.NavUp.Add_Click({ $i=$controls.NavList.SelectedIndex; if($i -gt 0){$tmp=$script:NavRows[$i];$script:NavRows[$i]=$script:NavRows[$i-1];$script:NavRows[$i-1]=$tmp;$controls.NavList.ItemsSource=$null;$controls.NavList.ItemsSource=$script:NavRows;$controls.NavList.SelectedIndex=$i-1} })
$controls.NavDown.Add_Click({ $i=$controls.NavList.SelectedIndex; if($i -ge 0 -and $i -lt $script:NavRows.Count-1){$tmp=$script:NavRows[$i];$script:NavRows[$i]=$script:NavRows[$i+1];$script:NavRows[$i+1]=$tmp;$controls.NavList.ItemsSource=$null;$controls.NavList.ItemsSource=$script:NavRows;$controls.NavList.SelectedIndex=$i+1} })
$controls.NavSave.Add_Click({ try { $rename=@{}; foreach($row in $script:NavRows){if($row.Key -ne $row.Label){$rename[$row.Key]=$row.Label}}; Set-MenuOrder -Root $Root -Order @($script:NavRows | ForEach-Object Key) -Rename $rename | Out-Null; Set-Status '导航设置已保存，发布后生效' } catch { [Windows.MessageBox]::Show($_.Exception.Message, '保存失败') | Out-Null } })

Refresh-Content; Load-Media; Load-Settings; Load-Navigation; Show-View 'Dashboard'
if ($SmokeTest) {
    Run-UiSmokeTest
    return
}
[void]$window.ShowDialog()
