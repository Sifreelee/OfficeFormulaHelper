## 排版公式与源码的区别

需要粘贴后显示成分式、上下标，并在公式编辑器中修改：先在下拉框明确选择“WPS 文字”“Word”“PowerPoint”或“WPS 演示”，点击“复制可编辑公式”，再回到目标文件按 Ctrl+V。WPS/Word 请用 .docx，演示请用 .pptx，且不要选择“粘贴为图片”或“仅保留文本”。粘贴完成前保持助手运行。

“复制 LaTeX 源码”仅复制代码文本，粘到普通正文不会自动切换公式模式。WPS 的“插入公式”现在也按所选 WPS 类型准备排版公式。

已在独立 WPS .docx 中验证：粘贴后为 OMath 数学公式、图片数量为零；修改字符、保存并重新打开后仍可编辑。旧 .doc 格式不保证保留原生数学公式。

# OfficeFormulaHelper v1.0.0

通用公式助手：从 GPT/Codex 复制内容中提取 LaTeX，支持源码复制、公式预览与 Office/Visio 输出。适用于 Windows 64 位桌面环境。

> 本软件免费使用，禁止售卖。若涉及侵权，请联系维护者处理。第三方组件遵循各自许可证。

## 使用方法

1. 下载完整工具包并解压，保留整个文件夹。
2. 双击“启动公式助手.cmd”或 OfficeFormulaHelper.exe。
3. 在 GPT/Codex 中选中公式或含公式的说明文字，按 Ctrl+C。
4. 助手自动过滤说明文字并提取 LaTeX。点击“复制 LaTeX 源码”，在 WPS、Word 或 PowerPoint 中按 Ctrl+V，即得到可以直接修改的代码文本。

快捷键 Ctrl+Alt+L 用于手动提取；Ctrl+Alt+V 用于在受支持的软件中调用相应输出。保持助手运行才能使用快捷键。自动提取可以关闭，以保留原始复制文字。

## 输出类型

| 功能 | 输出 |
| --- | --- |
| 复制 LaTeX 源码 | 纯文本代码，保留上下标、分式等 LaTeX 语法 |
| WPS 文字、WPS 演示模式 | 复制可编辑排版公式，使用 .docx/.pptx；源码文本通过独立按钮复制 |
| Word、PowerPoint 的“插入公式” | Office 原生排版公式 |
| Visio 的“插入公式” | SVG 矢量图片，可通过助手读取源码后更新 |

多条公式按原顺序提取并去重，合并为多行 LaTeX，行间增加空白。可通过“放大核对”同时检查预览和源码。提取依赖复制内容中保留的公式源码；对于只有截图或丢失结构的普通字符，不猜测上下标。

需要区分 LaTeX 源码文本和排版后的公式对象：源码文本在 .doc/.docx 中都可以编辑；现代排版公式建议使用 .docx/.pptx，旧格式可能降低为图片。

## 运行依赖

- Windows 64 位及 .NET Framework 4.x。
- vendor 内附 Node.js、MathJax 与 sharp；请勿只复制 EXE。
- 仅提取/复制 LaTeX 与图片预览，不需要启动 Microsoft Office 或 WPS。
- Word/PPT 原生排版输出需要本机安装相应软件及 Microsoft Office 的 MML2OMML.XSL 转换组件。

工具生成文件位于 output/ 和 temp/，已加入 Git 忽略规则。不会自动保存用户文档。

## 从源码构建

在 PowerShell 中运行：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\build.ps1
```

运行本地提取与纯文本输出检查：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\build.ps1 -Test
```

构建使用 Windows 自带的 .NET Framework C# 编译器，无需额外安装 SDK。版本号在 AssemblyInfo.cs 和 VERSION 中定义。

## 项目文件

- FormulaHelper.cs：桌面界面与 Office/Visio 入口。
- ClipboardLatex.cs：剪贴板公式源码提取、自动过滤文字和源码复制。
- NativeMath.cs：Office 数学格式转换与排版输出。
- render.cjs：MathJax 预览渲染。
- 使用说明.md：更详细的操作说明。

本仓库保留第三方依赖的许可证。v1.0.0 清理了重复版本、测试图片及临时产物；发布包只包含一个主程序。