# TeXShift 反向转换技术概述

## 概述

OneNote XML → Markdown 的反向转换功能，与正向转换架构对称设计。

**默认行为**：将选区/光标范围转换为 Markdown，以纯文本写回 OneNote（替换原内容）。

**设计范围**：
- **优先保证**：TeXShift 生成的 OneNote XML → Markdown 可逆
- **非 TeXShift 样式**：Best-effort 并输出 warning
- **识别依据**：配置值 + 结构特征（非硬编码颜色）

## 目录结构

```
src/TeXShift.Core/
├── OneNoteToMarkdown/                    # 反向转换模块
│   ├── Abstractions/
│   │   ├── IOneNoteToMarkdownConverter.cs
│   │   └── IOneNoteConverterContext.cs
│   ├── Handlers/
│   │   ├── IElementHandler.cs
│   │   ├── HeadingElementHandler.cs
│   │   ├── ParagraphElementHandler.cs
│   │   ├── ListElementHandler.cs
│   │   ├── CodeBlockElementHandler.cs
│   │   ├── QuoteBlockElementHandler.cs
│   │   ├── TableElementHandler.cs        # partial class
│   │   ├── TableElementHandler.Html.cs   # HTML 解析辅助
│   │   ├── MathElementHandler.cs
│   │   ├── MermaidElementHandler.cs
│   │   ├── HorizontalRuleElementHandler.cs
│   │   ├── ImageElementHandler.cs
│   │   └── OneNoteTableHelpers.cs
│   ├── Inlines/
│   │   ├── InlineParser.cs               # partial class 主入口
│   │   ├── InlineParser.Parse.cs         # 解析逻辑
│   │   ├── InlineParser.Html.cs          # HTML 标签处理
│   │   ├── InlineParser.Style.cs         # 样式属性解析
│   │   ├── InlineParser.Text.cs          # 文本处理
│   │   ├── InlineParser.InlineCode.cs    # 内联代码识别
│   │   ├── InlineParseMode.cs            # 解析模式枚举
│   │   └── HtmlStripper.cs               # 去除语法高亮 HTML
│   ├── OneNoteToMarkdownConverter.cs
│   ├── ReverseConversionResult.cs
│   └── HtmlRegexes.cs
├── OneNoteMeta/                          # 嵌入式元数据
│   ├── TeXShiftMetaKeys.cs
│   ├── TeXShiftMetaReader.cs
│   └── TeXShiftMetaWriter.cs
├── Services/
│   ├── ReverseConversion/
│   │   ├── ReverseSelectionPromoter.cs   # 选择升级
│   │   └── TeXShiftMetaTableSnippetRestorer.cs
│   ├── ReverseConversionOrchestrator.cs
│   ├── ReverseConversionOptions.cs
│   └── ReverseConversionPipelineResult.cs
└── Utils/
    ├── OneNoteHtmlEntityDecoder.cs
    └── OneNoteAutoLinkNormalizer.cs
```

## 架构设计

### 对称性映射

| 正向转换 | 反向转换 |
|---------|---------|
| `IMarkdownConverter` | `IOneNoteToMarkdownConverter` |
| `MarkdownToOneNoteConverter` | `OneNoteToMarkdownConverter` |
| `IBlockHandler` (处理 `Block`) | `IElementHandler` (处理 `XElement`) |
| `IMarkdownConverterContext` | `IOneNoteConverterContext` |
| `InlineRenderer` (Inline → HTML) | `InlineParser` (HTML → Markdown) |

### 管道流程

```
Read → Selection Promotion → Convert → Meta Table Restore → Write Back
```

**ReverseConversionOrchestrator** 编排：
1. **Read**: `OneNotePageReader.ExtractContentAsync()` 提取选区/光标范围内容
2. **Selection Promotion**: `ReverseSelectionPromoter` 升级选择范围
3. **Convert**: `OneNoteToMarkdownConverter.ConvertToMarkdownAsync()` XML → Markdown
4. **Meta Table Restore**: `TeXShiftMetaTableSnippetRestorer` 恢复被选中的表格片段
5. **Write Back**: `OneNotePageWriter.ReplaceContentAsync()` 写回纯文本

## 选择升级机制

**ReverseSelectionPromoter** 提供三种升级策略：

1. **整 Outline 升级**：选区覆盖整个 Outline + 有效 Meta → Cursor 模式（恢复源代码）
2. **表格升级**：选区在表格内部 → 升级到表格容器（避免留下空 OEChildren）
3. **嵌入对象升级**：光标在数学/图片等嵌入对象 + 有效 Meta → 升级到 Outline

**TeXShiftMetaTableSnippetRestorer**：当选择表格一部分时，用 Meta 中保存的原始表格片段替代通用渲染器输出，保证往返保真度。

## 元素识别策略

**核心原则**：
- **不硬编码字体/颜色**：样式参数必须来自设置（`OneNoteStyleConfig`），不可写死
- **颜色值不参与判别**：最多只把"是否存在底色属性"当作结构信号
- 标题 h1-h6 按设置字号做"就近反转"（nearest mapping）

### 识别开关

设置项 `TryRecognizeNonTeXShiftFormats`：

- **关闭（严格）**：仅识别 TeXShift 标准格式，误判最少
- **开启（默认）**：基于结构特征 + 字体/行距特征做兜底推断

### 元素识别规则

| 元素类型 | 识别特征 |
|---------|---------|
| 标题 | 字体大小 + spaceBefore/spaceAfter 匹配配置；兜底：>=14pt + 粗体 |
| 代码块 | 单列 Table + `hasHeaderRow="false"` + `shadingColor` 存在 + 代码字体/行距 |
| 引用块 | 单列 Table + `hasHeaderRow="false"` + `shadingColor` 存在 + 非代码特征 |
| 无序列表 | `OE/List/Bullet` |
| 有序列表 | `OE/List/Number` |
| 任务列表 | `OE/Tag[@completed]` |
| 表格 | `Table[@bordersVisible="true"]` |
| MathML | `T` 内容包含 `<!--[if mathML]>` |
| Mermaid | `Image[@alt="mermaid"]` |
| 水平线 | 居中 OE + (图片 <=4px 高 或 重复特殊字符) |
| 内联代码 | 代码字体 + `background-color` 属性存在 |
| 粗体/斜体/删除线/高亮/下划线/上下标 | CSS 样式属性及受支持的内联标签 |

### 表格处理

- **Header 推断**：
  - 严格：仅当 `hasHeaderRow="true"`
  - 兜底：首行全粗体也当 header
- **表头加粗**：不因"header 行整格加粗包装"输出 `**...**`（避免往返叠加）
- **对齐**：检查 OE `alignment` 属性 / CSS `text-align`

## 双通道策略

### 1. 源信息优先（Meta 通道）

若 Outline 携带 TeXShift 写入的 `<one:Meta>` 源信息，且签名一致，直接输出存储的 Markdown 源码。

### 2. 解析兜底（XML 解析通道）

若没有源信息或签名过期，解析 OneNote XML 生成语义等价的 Markdown。

## 嵌入式元数据方案

### Meta 键定义

```
texshift-schema       : 版本（"1"）
texshift-mode         : "render" / "source"
texshift-sourceEncoding: "plain-v1"
texshift-source-0...N : Markdown 源码分片（MaxChunkLength=8000）
texshift-sigVersion   : 签名算法版本（"1"）
texshift-sig          : 内容签名（SHA256）
```

### 签名算法

v1：以 `one:T` 的可见文本（剥离 HTML 标签、HTML 解码、归一化空白）+ MathML 内容哈希 + 图片 base64 哈希构造 token 流并签名。

### 过期判定

- `sigNow == texshift-sig`：未被修改 → 使用 Meta 源码
- `sigNow != texshift-sig`：已被修改 → 丢弃 Meta，回落到解析兜底

## 特殊元素处理

### 代码块

1. **HtmlStripper** 剥离语法高亮 `<span style='color:...'>`
2. 语言信息随 Meta 保留；解析兜底时省略语言标记

### 列表嵌套

根据 `<Indent>` 配置计算缩进级别，重建 Markdown 列表结构。

### 数学公式

Meta 优先恢复原始 LaTeX；签名不一致则输出占位符：
```
【TeXShift：公式已被修改，无法还原】
```

### Mermaid

Meta 优先恢复源码；源缺失则输出占位符。

### 图片

输出占位符（不内联 base64，避免性能问题）。

## 不可逆元素处理

- **输出占位符**：避免静默丢失内容
- **收集警告**：`ReverseConversionResult.Warnings`
- **结果呈现**：转换完成后向用户显示警告

## 错误处理

**降级策略**：
1. 主策略失败 → 尝试备用策略
2. 备用策略失败 → 输出占位符并记录 warning
3. 永不抛异常中断整个转换
