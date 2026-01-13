# TeXShift 反向转换实现计划

## 概述

实现 OneNote XML → Markdown 的反向转换功能，与现有正向转换架构对称设计。

## 架构设计

### 对称性映射

| 正向转换 | 反向转换 |
|---------|---------|
| `IMarkdownConverter` | `IOneNoteToMarkdownConverter` |
| `MarkdownToOneNoteConverter` | `OneNoteToMarkdownConverter` |
| `IBlockHandler` (处理 `Block`) | `IElementHandler` (处理 `XElement`) |
| `IMarkdownConverterContext` | `IOneNoteConverterContext` |
| `InlineRenderer` (Inline → HTML) | `InlineParser` (HTML → Markdown) |

### 新增目录结构

```
src/TeXShift.Core/
├── OneNoteToMarkdown/                    # 新增：反向转换模块
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
│   │   ├── TableElementHandler.cs
│   │   ├── MathElementHandler.cs
│   │   ├── MermaidElementHandler.cs
│   │   ├── HorizontalRuleElementHandler.cs
│   │   └── ImageElementHandler.cs
│   ├── Inlines/
│   │   └── InlineParser.cs
│   ├── OneNoteToMarkdownConverter.cs
│   └── ReverseConversionResult.cs
├── Metadata/                             # 新增：元数据缓存
│   ├── IMetadataCache.cs
│   └── PageMetadataCache.cs
└── Math/
    └── MathMLToLatexConverter.cs         # 新增：MathML反向转换
```

## 元素识别策略

| 元素类型 | 识别特征 |
|---------|---------|
| 标题 | `quickStyleIndex="0-5"` 或 字体>=14pt+粗体 |
| 代码块 | Table + `Cell[@shadingColor="#0D1117"]` |
| 引用块 | Table + `Cell[@shadingColor="#E8F5E9"]` |
| 无序列表 | `OE/List/Bullet` |
| 有序列表 | `OE/List/Number` |
| 任务列表 | `OE/Tag[@completed]` |
| 表格 | `Table[@bordersVisible="true"][@hasHeaderRow="true"]` |
| MathML | `T` 内容包含 `<!--[if mathML]>` |
| Mermaid | `Image[@alt="mermaid"]` |
| 水平线 | 居中OE + 特定图片/字符模式 |
| 内联代码 | `font-family:Consolas` + `background-color:#F1F1F1` |
| 粗体 | `font-weight:bold` |
| 斜体 | `font-style:italic` |
| 删除线 | `text-decoration:line-through` |

## 元数据缓存方案

### 目的
为无法完美还原的元素缓存原始数据：
- **Math**: 原始 LaTeX + 生成的 MathML 哈希（用于变更检测）
- **Mermaid**: 原始代码（PNG 不可逆）
- **图片**: 原始 URL/路径

### 存储位置
`%APPDATA%\TeXShift\metadata_cache.json`

### 缓存结构
```json
{
  "version": "1.0",
  "pages": {
    "{page-id}": {
      "math": {
        "{object-id}": {
          "latex": "...",
          "displayMode": true,
          "mathmlHash": "..."
        }
      },
      "mermaid": { "{object-id}": { "code": "..." } },
      "images": { "{object-id}": { "url": "...", "alt": "..." } }
    }
  }
}
```

### 缓存使用策略（用户确认）
1. **Math**: 对比当前 MathML 哈希与缓存哈希
   - 哈希匹配 → 使用缓存的 LaTeX
   - 哈希不匹配 → fallback 到 MathML 反向解析
2. **Mermaid**: 无法检测变更，直接使用缓存（如有）
3. **无缓存时**: 跳过元素，在结果中添加警告

### 集成方式
1. **正向转换时写入**: 修改 `MathBlockHandler`、`MermaidBlockHandler` 等
2. **反向转换时读取**: 各 `ElementHandler` 先查缓存，验证后使用

## MathML → LaTeX 反向转换

### 策略优先级
1. 查缓存获取原始 LaTeX（最可靠）
2. 尝试 MathML 反向解析（尽力而为）
3. 返回占位符注释（最后手段）

### MathML 反向解析映射
```
mfrac → \frac{num}{den}
msqrt → \sqrt{content}
msub → base_{sub}
msup → base^{sup}
mfenced → \left( ... \right)
mtable → \begin{matrix} ... \end{matrix}
```

## 实现阶段

### 阶段1：核心框架
- [ ] `IOneNoteToMarkdownConverter` 接口
- [ ] `IElementHandler` 接口
- [ ] `IOneNoteConverterContext` 接口
- [ ] `OneNoteToMarkdownConverter` 骨架
- [ ] `InlineParser` 基础实现

**验证**: 简单段落的往返测试

### 阶段2：基础元素处理器
- [ ] `ParagraphElementHandler`
- [ ] `HeadingElementHandler`
- [ ] `ListElementHandler`（ul/ol/task）

**验证**: 标题h1-h6、嵌套列表、任务列表的往返测试

### 阶段3：容器元素处理器
- [ ] `CodeBlockElementHandler`
- [ ] `QuoteBlockElementHandler`
- [ ] `TableElementHandler`
- [ ] `HorizontalRuleElementHandler`

**验证**: 代码块、嵌套引用、表格对齐的往返测试

### 阶段4：元数据缓存
- [ ] `IMetadataCache` 接口
- [ ] `PageMetadataCache` 实现
- [ ] 修改正向转换 Handler 写入缓存
- [ ] 反向 Handler 集成缓存读取

**验证**: 缓存读写单元测试

### 阶段5：特殊元素处理器
- [ ] `MathElementHandler`（缓存 + MathML 反向）
- [ ] `MermaidElementHandler`（仅缓存）
- [ ] `ImageElementHandler`
- [ ] `MathMLToLatexConverter`

**验证**: 数学公式、Mermaid 图的往返测试

### 阶段6：集成与UI
- [ ] `ServiceContainer` 添加工厂方法
- [ ] `ConversionOrchestrator` 添加反向转换支持
- [ ] Ribbon 添加"反向转换"按钮（先复制到剪贴板，导出方式后续再定）
- [ ] E2E 测试

**验证**: 完整 E2E 往返测试

## 关键文件

### 需要创建
| 文件 | 用途 |
|-----|-----|
| `OneNoteToMarkdown/Abstractions/IOneNoteToMarkdownConverter.cs` | 主接口 |
| `OneNoteToMarkdown/OneNoteToMarkdownConverter.cs` | 核心实现 |
| `OneNoteToMarkdown/Handlers/*.cs` | 各元素处理器 |
| `OneNoteToMarkdown/Inlines/InlineParser.cs` | HTML→MD 解析 |
| `Metadata/PageMetadataCache.cs` | 缓存实现 |
| `Math/MathMLToLatexConverter.cs` | MathML 反向 |

### 需要修改
| 文件 | 修改内容 |
|-----|---------|
| `Services/ServiceContainer.cs` | 添加反向转换器工厂 |
| `Markdown/Handlers/MathBlockHandler.cs` | 写入缓存 |
| `Markdown/Handlers/MermaidBlockHandler.cs` | 写入缓存 |
| `Services/ConversionOrchestrator.cs` | 添加反向转换选项 |

## 不可逆元素处理（用户确认）

对于无法还原的元素（无缓存的 Mermaid、无法解析的 MathML 等）：
- **跳过该元素**：不输出任何内容
- **收集警告**：在 `ReverseConversionResult.Warnings` 中记录
- **结果呈现**：转换完成后向用户显示警告列表

## 错误处理

### 新增异常类型
```csharp
public class ReverseConversionException : TeXShiftException
{
    // Error code: TSE006
}
```

### 降级策略
1. 主策略失败 → 尝试备用策略
2. 备用策略失败 → 生成占位符注释
3. 永不抛异常中断整个转换

## 测试策略

### 单元测试
- `InlineParser` 各种 HTML 样式解析
- 各 `ElementHandler` 独立测试
- `MathMLToLatexConverter` 各种数学结构

### 集成测试
- 元数据缓存读写
- 正向→反向往返保真度

### E2E 测试
- 完整文档往返
- 包含所有元素类型的综合测试
