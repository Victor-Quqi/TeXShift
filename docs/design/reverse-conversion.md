# TeXShift 反向转换实现计划

## 概述

实现 OneNote XML → Markdown 的反向转换功能，与现有正向转换架构对称设计。

默认行为：将选区/光标范围转换为 Markdown，并以纯文本写回 OneNote（替换原内容；可选复制到剪贴板）。

## 设计范围

**明确 Scope**：
- **优先保证**：TeXShift 生成的 OneNote XML → Markdown 可逆
- **非 TeXShift 样式**：Best-effort 并输出 warning
- **输出**：将 Markdown 结果作为纯文本写回 OneNote（替换原内容；可选复制到剪贴板）
- 识别依据：配置值 + 结构特征（非硬编码颜色）

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
│   │   ├── InlineParser.cs
│   │   └── HtmlStripper.cs              # 去除语法高亮HTML
│   ├── OneNoteToMarkdownConverter.cs
│   └── ReverseConversionResult.cs
├── Metadata/                             # 新增：元数据缓存
│   ├── IMetadataCache.cs
│   └── ContentHashMetadataCache.cs      # 基于内容哈希的缓存
└── Math/
    └── MathMLToLatexConverter.cs         # 新增：MathML反向转换
```

## 元素识别策略

**核心原则**：从 `OneNoteStyleConfig` 读取配置值，而非硬编码颜色

| 元素类型 | 识别特征 |
|---------|---------|
| 标题 | `quickStyleIndex="0-5"` 或 字体>=14pt+粗体 |
| 代码块 | 单列Table + `hasHeaderRow="false"` + Cell背景色匹配 `CodeBlockConfig.BackgroundColor` + OE样式含 Consolas |
| 引用块 | 单列Table + `hasHeaderRow="false"` + Cell背景色匹配 `QuoteBlockConfig.BackgroundColor` + 子内容非代码风格 |
| 无序列表 | `OE/List/Bullet` |
| 有序列表 | `OE/List/Number` |
| 任务列表 | `OE/Tag[@completed]` |
| 表格 | `Table[@bordersVisible="true"][@hasHeaderRow="true"]` |
| MathML | `T` 内容包含 `<!--[if mathML]>` |
| Mermaid | `Image[@alt="mermaid"][@format="png"]` |
| 水平线 | 居中OE + (图片2px高度 或 重复特殊字符) |
| 内联代码 | `font-family:Consolas` + `background-color` 匹配 `InlineCodeConfig` |
| 粗体 | `font-weight:bold` |
| 斜体 | `font-style:italic` |
| 删除线 | `text-decoration:line-through` |

## 元数据缓存方案（改进版）

### 缓存键策略

**用内容哈希作为主键**（非 objectID）：
- objectID 在 `UpdatePageContent` 后由 OneNote 分配，Handler 阶段拿不到
- 内容哈希可跨复制/粘贴/重排保持关联

### 缓存结构
```json
{
  "version": "1.0",
  "entries": {
    "math": {
      "{normalized-mathml-hash}": {
        "latex": "\\frac{a}{b}",
        "displayMode": true
      }
    },
    "mermaid": {
      "{image-bytes-hash}": {
        "code": "graph TD\\n  A-->B",
        "language": "mermaid"
      }
    },
    "codeblock": {
      "{content-hash}": {
        "language": "python"
      }
    }
  }
}
```

### 缓存使用策略（用户确认）
1. **Math**:
   - 计算当前 MathML 规范化后的哈希
   - 哈希命中 → 使用缓存的 LaTeX
   - 哈希不命中 → fallback 到 MathML 反向解析
2. **Mermaid**: 计算图片字节哈希，查缓存
3. **CodeBlock**: 缓存语言信息（fenced.Info 正向时丢失）
4. **无缓存时**: 跳过元素，在结果中添加警告

### 存储位置
`%APPDATA%\TeXShift\metadata_cache.json`

### 集成方式
1. **正向转换时写入**: 修改 `MathBlockHandler`、`MermaidBlockHandler`、`CodeBlockHandler`
2. **反向转换时读取**: 各 `ElementHandler` 先查缓存，验证后使用

## 代码块反向转换难点

### 问题
1. 语法高亮输出富 HTML，需要"去高亮"
2. 语言信息 `fenced.Info` 未写入 OneNote XML，反向无法恢复

### 解决方案
1. **HtmlStripper**: 专门的去高亮工具，剥离所有 `<span style='color:...'>` 保留纯文本
2. **语言缓存**: 正向转换时缓存 `{内容哈希: language}`，反向时查缓存恢复

## 列表树结构处理

### 问题
正向把列表挂到"前一个容器块"的 OEChildren 下（`MarkdownToOneNoteConverter.cs:171`），
反向若按 XML 层级直接遍历，输出顺序/缩进会错乱。

### 解决方案
**反嵌套/反扁平化策略**：
1. 识别 `OEChildren` 中的列表项
2. 根据 `<Indent>` 配置计算缩进级别
3. 重建 Markdown 列表结构

## MathML → LaTeX 反向转换

### 策略优先级
1. 查缓存获取原始 LaTeX（最可靠）
2. 检查 MathML `<annotation encoding="application/x-tex">` 标签（如果 OneNote 保留）
3. 尝试 MathML 反向解析（尽力而为）
4. 跳过并警告（最后手段）

### 备选方案
**正向转换时写入 annotation**：
```xml
<mml:semantics>
  <mml:mrow>...</mml:mrow>
  <mml:annotation encoding="application/x-tex">\frac{a}{b}</mml:annotation>
</mml:semantics>
```
若 OneNote 保留此标签，可减少外部缓存依赖。需要测试验证。

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
- [ ] `IOneNoteConverterContext` 接口（注入 `OneNoteStyleConfig` 用于配置值识别）
- [ ] `OneNoteToMarkdownConverter` 骨架
- [ ] `InlineParser` 基础实现（粗体、斜体、删除线、链接）
- [ ] `HtmlStripper` 去高亮工具
- [ ] `TeXShift.Tests.E2E` 增加 `reverse-xml` / `reverse-inplace`（用于阶段性验证与写回验证）

**验证**: `reverse-xml` 验证 XML→Markdown（无需 Ribbon）

### 阶段2：基础元素处理器
- [ ] `ParagraphElementHandler`
- [ ] `HeadingElementHandler`
- [ ] `ListElementHandler`（ul/ol/task，含反嵌套逻辑）

**验证**: 标题h1-h6、嵌套列表、任务列表的往返测试

### 阶段3：容器元素处理器
- [ ] `CodeBlockElementHandler`（含去高亮，暂无语言恢复）
- [ ] `QuoteBlockElementHandler`
- [ ] `TableElementHandler`
- [ ] `HorizontalRuleElementHandler`

**验证**: 代码块去高亮、嵌套引用、表格对齐的往返测试
**里程碑**: 阶段1-3完成后，基础反向转换可用

### 阶段4：元数据缓存
- [ ] `IMetadataCache` 接口
- [ ] `ContentHashMetadataCache` 实现（基于内容哈希）
- [ ] 修改正向转换 Handler 写入缓存（Math/Mermaid/CodeBlock语言）
- [ ] 反向 Handler 集成缓存读取

**验证**: 缓存读写单元测试、语言恢复测试

### 阶段5：特殊元素处理器
- [ ] `MathElementHandler`（缓存 + annotation + MathML 反向）
- [ ] `MermaidElementHandler`（仅缓存）
- [ ] `ImageElementHandler`
- [ ] `MathMLToLatexConverter`
- [ ] 测试 annotation 标签是否被 OneNote 保留

**验证**: 数学公式、Mermaid 图的往返测试

### 阶段6：集成与UI
- [ ] `ServiceContainer` 添加工厂方法
- [ ] `ConversionOrchestrator` 添加反向转换支持
- [ ] Ribbon 添加"反向转换"按钮（将选区/光标范围转换为 Markdown，并以纯文本写回 OneNote；可选复制到剪贴板）
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
| `OneNoteToMarkdown/Inlines/HtmlStripper.cs` | 去高亮工具 |
| `Metadata/ContentHashMetadataCache.cs` | 基于内容哈希的缓存 |
| `Math/MathMLToLatexConverter.cs` | MathML 反向 |

### 需要修改
| 文件 | 修改内容 |
|-----|---------|
| `Services/ServiceContainer.cs` | 添加反向转换器工厂 |
| `Markdown/Handlers/MathBlockHandler.cs` | 写入缓存 |
| `Markdown/Handlers/MermaidBlockHandler.cs` | 写入缓存 |
| `Markdown/Handlers/CodeBlockHandler.cs` | 写入语言缓存 |
| `Math/OneNoteMathMLAdapter.cs` | 可选：添加 annotation 标签 |
| `Services/ConversionOrchestrator.cs` | 添加反向转换选项 |
| `tests/TeXShift.Tests.E2E/Program.cs` | 增加 `reverse-xml` 子命令 |

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
2. 备用策略失败 → 跳过并记录警告
3. 永不抛异常中断整个转换

## 测试策略

### 单元测试
- `InlineParser` 各种 HTML 样式解析
- `HtmlStripper` 去高亮测试
- 各 `ElementHandler` 独立测试
- `MathMLToLatexConverter` 各种数学结构
- `ContentHashMetadataCache` 哈希一致性

### 集成测试
- 元数据缓存读写
- 正向→反向往返保真度

### E2E 测试
- 完整文档往返
- 包含所有元素类型的综合测试
- 无需 Ribbon：E2E 增加 `reverse-xml`（读取 03/04 XML，输出 Markdown）
- 无需 Ribbon：E2E 增加 `reverse-inplace`（基于 OneNote COM：选区 XML → Markdown → 写回 OneNote）


