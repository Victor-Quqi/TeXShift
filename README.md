# TeXShift: Bridging OneNote and Plain Text

English | [简体中文](README.zh-CN.md)

**TeXShift** is a COM Add-in for Microsoft OneNote designed to address the core pain points that engineers, researchers, and students face when taking technical notes in OneNote: **building a bridge between a powerful rich text editor and efficient plain text markup languages like Markdown and LaTeX.**

---

## Features

### Implemented

- **Markdown Conversion** - Convert Markdown syntax to OneNote rich text
  - Headings (H1-H6)
  - Ordered/unordered lists
  - Task lists (checkboxes)
  - Block quotes (nested support)
  - Tables (header bold, column alignment)
  - Links
  - Embedded images
  - Horizontal rules

- **LaTeX Formulas** - Convert LaTeX math expressions to native editable OneNote equations
  - Built-in MathJax resources for fully offline operation
  - Requires .NET Framework 4.8 and WebView2 Runtime

- **Syntax Highlighting** - ColorCode-based code highlighting
  - Transform code blocks into rich text with background color and syntax coloring

- **Intelligent Selection** - Two operation modes
  - **Cursor Mode**: Operates on the entire text box
  - **Selection Mode**: Operates on complete paragraphs containing highlighted text

- **Fully Offline** - All core features work locally without network connection

### In Development

- Live preview pane
- Mermaid diagram support
- Custom style settings
- Rich text to Markdown (reverse conversion)

## Tech Stack

- **Language:** C# (.NET Framework 4.8)
- **Framework:** OneNote COM Add-in
- **IDE:** Visual Studio 2026
- **Dependencies:**
  - Markdig - Markdown parsing
  - ColorCode - Syntax highlighting
  - MathJax - LaTeX rendering

## System Requirements

- Windows 10/11
- Microsoft OneNote Desktop (x64)
- .NET Framework 4.8
- WebView2 Runtime (for LaTeX rendering)

## Installation

The project is still under active development. No stable release yet.

## License

This project is licensed under [GPLv3](LICENSE).
