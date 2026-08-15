# ZZ-100.github.io

张振的个人学术主页，使用 Hexo 生成并通过 GitHub Actions 部署到 GitHub Pages。

## 本地使用

首次安装依赖：

```powershell
npm ci
```

本地预览：

```powershell
npm run server
```

构建检查：

```powershell
npm run clean
npm run build
```

Windows 用户运行 `start.cmd` 即可打开新版 WPF 图形化管理工具。工具按“概览、内容、媒体、导航、设置、发布”分区，支持内容搜索/筛选、新建文章和独立页面、Word 往返编辑与草稿重置、网页预览、头像与 CV 管理、导航排序、站点设置、可复制的发布检查结果、构建检查和 Git 发布。

如需检查界面控件、最小窗口布局和头像文件锁定状态，可运行内置自检：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -STA -File .\blog-gui-wpf.ps1 -SmokeTest
```

如果需要使用旧版 WinForms 界面，可手动运行：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -STA -File .\blog-gui.ps1
```

## 内容位置

- `source/_posts/home.md`：主页内容
- `source/about/index.md`：About 页面
- `themes/Academia/_config.yml`：作者简介、导航和主题设置
- `themes/Academia/source/img/profile.png`：头像

没有真实论文或个人主页时，不要保留平台首页、测试 CV 或“待补充”占位链接；暂时没有内容的栏目应直接隐藏。

## Word 编辑约定

Word 编辑是 Markdown 的辅助入口，不是第二份长期内容源。当前转换稳定支持标题、段落、项目符号、粗体和斜体；表格、图片、复杂列表、代码块和复杂链接应在 Markdown 中复核。工具会记录 Word 文件哈希，未导入的修改不会被静默覆盖。

## 发布流程

图形化工具和命令行工具都会先执行清理、Hexo 构建和内部链接检查，再提交当前分支并推送。GitHub Actions 监听 `main` 分支，构建 `public` 后部署到 GitHub Pages。

删除内容会先移动到项目内的 `.trash/` 回收目录；该目录不会被提交。确认无误后可手动清理回收目录。
