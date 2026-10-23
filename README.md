# DocPreview

一个只读的 `.docx` 阅读器，带暗色模式。Word / WPS 写好文档后，用它舒服地读。

**完全不修改、不依赖 Word**：`.docx` 本质是 zip 包，本程序自己解析其中的 XML，用系统自带的 WebView2 渲染。

## 下载

到 [Releases](https://github.com/xsnj-9527/rea-word/releases) 页面下载：

| 文件 | 说明 |
|---|---|
| `DocPreview-x.y.z-setup.exe` | 安装包，双击安装，有开始菜单和卸载项 |
| `DocPreview-x.y.z-portable.zip` | 绿色版，解压即用，不写注册表 |

安装包安装时会问你**是否把 DocPreview 加入 .docx 的右键「打开方式」**（默认勾选）。
它**不会**修改你的默认打开程序——双击 .docx 还是 WPS/Word 打开。

## 使用

- **右键打开**：`.docx` 文件右键 → 打开方式 → DocPreview
- **拖放**：把文件拖进窗口
- **命令行**：`DocPreview.exe "文档.docx"`

## 功能

- 暗色 / 亮色主题（默认暗色，会记住选择）
- **内嵌图片显示**，点击图片全屏放大，`Esc` 关闭
- **文字可选中复制**（鼠标选择 + `Ctrl+C`）
- 标题层级、加粗、斜体、下划线、删除线、上下标
- 字体、字号、颜色、段落对齐、缩进
- 表格（含合并单元格、嵌套表格）
- 缩放：`Ctrl` + `+` / `-` / `0`
- 打开：`Ctrl+O`
- 状态栏显示段落数、字数，以及不支持内容的提示

### 暗色模式下的黑色文字

Word 允许作者把文字设成黑色。深色页面上这会变成「黑底黑字」，所以程序会自动
把**暗且接近灰色**的文字翻转成浅色，同时保留红色、蓝色等强调色。

## 已知限制

- **只读**，不能编辑
- 排版是阅读版式，不是 Word 的精确分页
- 不支持：文本框内容、图表、SmartArt、公式、EMF/WMF 矢量图
- 修订标记按「接受后」状态显示
- 遇到上述内容会在状态栏提示，不会静默丢弃

## 从源码运行

```bash
pip install pywebview
python src/app.py "文档.docx"
```

需要 Python 3.11+ 和 Windows 10/11（依赖系统自带的 WebView2）。

## 构建

```bash
build.bat
```

会依次：用 PyInstaller 打包 `dist\DocPreview.exe` → 用 Inno Setup 生成安装包
→ 生成绿色版 zip。产物都在 `dist\` 和 `installer_out\`。

依赖：[PyInstaller](https://pyinstaller.org/) 和 [Inno Setup 6](https://jrsoftware.org/isinfo.php)。

## 更新日志

见 [CHANGELOG.md](CHANGELOG.md)。

## 许可

[MIT](LICENSE)
