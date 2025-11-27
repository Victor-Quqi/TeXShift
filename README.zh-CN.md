# TeXShift: 连接 OneNote 与纯文本标记语言

[English](README.md) | 简体中文

**TeXShift** 是一个为 Microsoft OneNote 开发的 COM 插件，致力于解决工程师、研究者和学生在 OneNote 中进行技术笔记记录时的核心痛点：**在强大的富文本编辑器与高效的纯文本标记语言（如 Markdown、LaTeX）之间建立一座桥梁。**

---

## 核心功能

### 已实现

- **Markdown 转换** - 将 Markdown 语法转换为 OneNote 富文本格式
  - 标题（H1-H6）
  - 有序/无序列表
  - 任务列表（复选框）
  - 引用块（支持嵌套）
  - 表格（支持表头加粗、列对齐）
  - 链接
  - 图片嵌入
  - 分割线

- **LaTeX 公式** - 将 LaTeX 数学公式转换为 OneNote 原生可编辑公式
  - 内置 MathJax 资源，完全离线运行
  - 需要 .NET Framework 4.8 和 WebView2 Runtime

- **代码高亮** - 基于 ColorCode 的语法高亮
  - 将代码块转换为带背景色和语法着色的富文本

- **智能选择** - 识别两种操作模式
  - **光标模式**：操作整个文本框
  - **选区模式**：操作高亮文字所在的完整段落

- **完全离线** - 所有核心功能均可在本地离线运行

### 开发中

- 实时预览窗格
- Mermaid 图表支持
- 自定义样式设置
- 富文本转 Markdown（反向转换）

## 技术栈

- **语言:** C# (.NET Framework 4.8)
- **框架:** OneNote COM Add-in
- **IDE:** Visual Studio 2026
- **依赖:**
  - Markdig - Markdown 解析
  - ColorCode - 语法高亮
  - MathJax - LaTeX 渲染

## 系统要求

- Windows 10/11
- Microsoft OneNote 桌面版 (x64)
- .NET Framework 4.8
- WebView2 Runtime（用于 LaTeX 渲染）

## 安装

项目仍在开发中，暂未发布稳定版本。

## 许可证

本项目采用 [GPLv3](LICENSE) 发布。
