# TeXShift: Bridging OneNote and Plain Text

**TeXShift** 是一个为 Microsoft OneNote 开发的 COM 插件，致力于解决工程师、研究者和学生在 OneNote 中进行技术笔记记录时的核心痛点：**在强大的富文本编辑器与高效的纯文本标记语言（如 Markdown, LaTeX）之间建立一座桥梁。**

---

## 核心功能 ✨

*   **✍️ 双向转换 (Bi-directional Conversion):**
    一键在 OneNote 原生格式（包括 OMML 公式、富文本样式）与纯文本标记语言之间进行切换。智能识别选区内容，自动判断转换方向。

*   **⚡ 实时预览 (Live Preview Pane):**
    提供一个可停靠的预览窗格，当您在 OneNote 中编辑 Markdown 或 LaTeX 时，实时渲染出最终的富文本效果，所见即所得。

*   **🖥️ 代码高亮 (Syntax Highlighting):**
    将 Markdown/LaTeX 中的代码块 (` ``` `) 优雅地转换为带有背景色和语法高亮的 OneNote 富文本块，让代码笔记清晰易读。

*   **📊 图表生成 (Diagrams as Code):**
    支持 [Mermaid](https://mermaid-js.github.io/mermaid/#/) 语法，将描述图表的文本一键转换为 PNG 图像并插入页面，实现“代码即图表”。

*   **🎯 智能选择 (Intelligent Selection):**
    识别 **光标模式**（操作整个文本框）与 **选区模式**（操作高亮文字所在的段落）。在选区模式下，插件会将包含选区的**整个段落**作为操作对象，确保转换的上下文完整性。

*   **🔧 高度可定制 (Highly Customizable):**
    通过设置界面可以自定义 Markdown 样式映射规则（如标题字号、代码块背景色）和常用功能的快捷键。

*   **🔌 完全离线 (Offline First):**
    所有核心功能均可在本地离线运行，无需网络连接，确保数据安全与随时随地的可用性。

## 技术栈 🛠️

*   **语言:** C# (.NET Framework)
*   **框架:** OneNote COM Add-in
*   **IDE:** Visual Studio 2026

## 项目状态 🚀

项目正处于积极开发阶段。开发优先级将首先聚焦于实现一个包含**核心转换逻辑**、**实时预览窗格**和**代码块高亮**的最小可行产品 (MVP)，以确保在课程项目展示中获得最佳效果。

## 安装与使用

项目仍在开发中，暂未发布稳定版本。

## 许可证

本项目采用 [GPLv3](LICENSE) 发布。