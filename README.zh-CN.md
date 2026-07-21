# TeXShift: 连接 OneNote 与纯文本标记语言

[English](README.md) | 简体中文

**TeXShift** 是一个为 Microsoft OneNote 开发的 COM 插件，致力于解决工程师、研究者和学生在 OneNote 中进行技术笔记记录时的核心痛点：**在强大的富文本编辑器与高效的纯文本标记语言（如 Markdown、LaTeX）之间建立一座桥梁。**

---

## 核心功能

- **Markdown 转换** - 将 Markdown 语法转换为 OneNote 富文本格式
  - 标题（H1-H6）
  - 有序/无序列表
  - 任务列表（复选框）
  - 引用块（支持嵌套）
  - 表格（支持表头加粗、列对齐）
  - 链接
  - 图片嵌入
  - 分割线（图片或字符样式）

- **LaTeX 公式** - 将 LaTeX 数学公式转换为 OneNote 原生可编辑公式
  - 内置 MathJax 资源，完全离线运行
  - 支持行内公式（`$...$`）和块级公式（`$$...$$`）

- **Mermaid 图表** - 将 Mermaid 图表渲染为嵌入式 PNG 图片
  - 支持流程图、时序图、类图等
  - 可配置主题和分辨率

- **代码高亮** - 基于 TextMateSharp 的语法高亮
  - 广泛的语言支持，精确的词法分析
  - 可自定义背景色、文字颜色、字体和行距

- **智能选择** - 识别两种操作模式
  - **光标模式**：操作整个文本框
  - **选区模式**：操作高亮文字所在的完整段落

- **自定义样式** - 通过设置界面完全控制外观
  - 引用块背景色
  - 标题字号（H1-H6）
  - 代码块样式（颜色、字体、字号、行距）
  - 行内代码样式
  - Mermaid 主题和最大分辨率
  - 分隔线样式（图片/字符）

- **反向转换** - 将 OneNote 富文本转换回 Markdown
  - 双通道策略：嵌入式元数据（无损）+ XML 解析（兜底）
  - 标题、段落、列表（有序/无序/任务）、表格
  - 代码块（自动剥离语法高亮）
  - 数学公式（通过元数据恢复原始 LaTeX）
  - Mermaid 图表（通过元数据恢复原始源码）
  - 内联样式（粗体、斜体、删除线、行内代码、链接）

- **本地化** - 支持中英文界面

- **完全离线** - 所有核心功能均可在本地离线运行

### 开发中

- 实时预览窗格

## 技术栈

- **语言:** C# (.NET Framework 4.8)
- **框架:** OneNote COM Add-in
- **UI:** WPF + Material Design
- **依赖:**
  - Markdig - Markdown 解析
  - TextMateSharp - 语法高亮
  - MathJax - LaTeX 渲染
  - Mermaid.js - 图表渲染

## 系统要求

- Windows 10/11
- Microsoft OneNote 桌面版 (x64)
- .NET Framework 4.8
- WebView2 Runtime（用于 LaTeX 和 Mermaid 渲染）

## 安装

从 [Releases](https://github.com/Victor-Quqi/TeXShift/releases) 下载最新的 x64 Setup `.exe`。

从 TeXShift 0.2.x 或更早版本升级时，请先卸载旧版。

Setup `.exe` 安装的卸载程序可选择一并删除安装用户的 TeXShift 设置、缓存和默认位置调试日志。

## 从源码构建

对于想要从源码构建 TeXShift 的开发者：

1. **克隆仓库**
   ```bash
   git clone https://github.com/Victor-Quqi/TeXShift.git
   cd TeXShift
   ```

2. **构建与测试**

   ```powershell
   .\build.ps1 -Target Build -Configuration Debug
   .\build.ps1 -Target Test -Configuration Debug
   ```

   构建脚本会恢复固定版本的 MathJax 和 NuGet 依赖，无需 Visual Studio。

3. **首次注册 Debug 加载项**

   正常构建并关闭 OneNote，然后在管理员 PowerShell 中运行：

   ```powershell
   .\setup\Register-TeXShiftDev.ps1 -Action Register
   ```

   后续构建沿用同一路径，重启 OneNote 即会加载新 DLL。普通 PowerShell 可用 `-Action Status` 检查状态；管理员 PowerShell 可用 `-Action Unregister` 移除开发版 Ribbon。

## 许可证

本项目采用 [GPLv3](LICENSE) 发布。
