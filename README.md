# TeXShift: Bridging OneNote and Plain Text

English | [简体中文](README.zh-CN.md)

**TeXShift** is a COM Add-in for Microsoft OneNote designed to address the core pain points that engineers, researchers, and students face when taking technical notes in OneNote: **building a bridge between a powerful rich text editor and efficient plain text markup languages like Markdown and LaTeX.**

---

## Features

- **Markdown Conversion** - Convert Markdown syntax to OneNote rich text
  - Headings (H1-H6)
  - Ordered/unordered lists
  - Task lists (checkboxes)
  - Block quotes (nested support)
  - Tables (header bold, column alignment)
  - Links
  - Embedded images
  - Horizontal rules (image or character style)

- **LaTeX Formulas** - Convert LaTeX math expressions to native editable OneNote equations
  - Built-in MathJax resources for fully offline operation
  - Inline (`$...$`) and display (`$$...$$`) math support

- **Mermaid Diagrams** - Render Mermaid diagrams as embedded PNG images
  - Flowcharts, sequence diagrams, class diagrams, and more
  - Configurable theme and resolution

- **Syntax Highlighting** - TextMateSharp-based code highlighting
  - Wide language support with accurate tokenization
  - Customizable background color, text color, font, and spacing

- **Intelligent Selection** - Two operation modes
  - **Cursor Mode**: Operates on the entire text box
  - **Selection Mode**: Operates on complete paragraphs containing highlighted text

- **Custom Styles** - Full control over appearance via settings UI
  - Quote block background color
  - Heading font sizes (H1-H6)
  - Code block styling (colors, font, size, line spacing)
  - Inline code styling
  - Mermaid theme and max resolution
  - Horizontal rule style (image/character)

- **Reverse Conversion** - Convert OneNote rich text back to Markdown
  - Dual-channel strategy: embedded metadata (lossless) + XML parsing (fallback)
  - Headings, paragraphs, lists (ordered/unordered/task), tables
  - Code blocks with syntax highlighting stripped
  - Math formulas (restore original LaTeX via metadata)
  - Mermaid diagrams (restore original source via metadata)
  - Inline styles (bold, italic, strikethrough, code, links)

- **Localization** - UI available in English and Simplified Chinese

- **Fully Offline** - All core features work locally without network connection

### In Development

- Live preview pane

## Tech Stack

- **Language:** C# (.NET Framework 4.8)
- **Framework:** OneNote COM Add-in
- **UI:** WPF with Material Design
- **Dependencies:**
  - Markdig - Markdown parsing
  - TextMateSharp - Syntax highlighting
  - MathJax - LaTeX rendering
  - Mermaid.js - Diagram rendering

## System Requirements

- Windows 10/11
- Microsoft OneNote Desktop (x64)
- .NET Framework 4.8
- WebView2 Runtime (for LaTeX and Mermaid rendering)

## Installation

Download the latest `.msi` installer from the [Releases](https://github.com/Victor-Quqi/TeXShift/releases) page and run it.

## Building from Source

For developers who want to build TeXShift from source:

1. **Clone the repository**
   ```bash
   git clone https://github.com/Victor-Quqi/TeXShift.git
   cd TeXShift
   ```

2. **Setup MathJax** (first time only)
   ```powershell
   .\setup-mathjax.ps1
   ```
   This downloads MathJax from npm (~23 MB). Requires Node.js and npm.

3. **Build the project**
   - **Full solution** (Add-in): Open in Visual Studio → Build → Rebuild Solution (Debug|x64)
   - **Core + Tests only**:
     ```powershell
     .\build.ps1 -Target Build -Configuration Debug
     ```

## License

This project is licensed under [GPLv3](LICENSE).
