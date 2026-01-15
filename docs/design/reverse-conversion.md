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
├── OneNoteMeta/                          # 新增：嵌入式元数据（<one:Meta>）读写与签名
│   ├── TeXShiftMetaKeys.cs
│   ├── TeXShiftMetaReader.cs
│   └── TeXShiftMetaWriter.cs
└── (其它模块不变)
```

## 元素识别策略

**核心原则**：从 `OneNoteStyleConfig` 读取配置值，而非硬编码颜色

## 总体策略（面向“实时预览/交换”）

反向转换分两条通道：**源信息优先**，**解析兜底**。

1. **源信息优先（TeXShift 管理的内容）**：
   - 若目标 `Outline`（或 `OE`）携带 TeXShift 写入的 `<one:Meta>` 源信息，且判定未过期，则直接输出其中保存的（可规范化的）Markdown 源码。
   - 该通道用于支持后续的“实时预览（source/render 双视图）”与“一键交换”。
2. **解析兜底（用户自写/源信息过期）**：
   - 若没有源信息，或源信息判定过期，则退化为解析 OneNote XML，并按识别规则生成语义等价的规范化 Markdown。

### 识别模式（严格 / 模糊）

为避免“用户自写内容被误判为 TeXShift 元素”，反向转换提供一个全局开关：

- **严格识别（关闭模糊识别）**：只在明确命中 TeXShift 正向转换的标记/样式时才还原为对应 Markdown（例如 `quickStyleIndex`、按当前配置生成的 `shadingColor` / `font-family` 等）。未命中则按普通文本处理，**误判更少**。
- **模糊识别（开启模糊识别，默认）**：在严格匹配失败时，允许用 **结构特征/字体特征** 做兜底（例如：单列表格 + 等宽字体更可能是代码块；单列表格 + 底色且不像代码更可能是引用块）。**不做“相近颜色”判断**（只做必要的格式归一化，如忽略 ARGB 的 alpha）。

该开关会放在设置里，旁边提供说明（悬停/点击查看），后续也可复用于其它设置项的“解释说明”。

| 元素类型 | 识别特征 |
|---------|---------|
| 标题 | `quickStyleIndex="0-5"` 或 字体>=14pt+粗体 |
| 代码块 | **严格**：单列Table + `hasHeaderRow="false"` + `shadingColor == CodeBlockConfig.BackgroundColor`（忽略 alpha） + 行 OE 命中代码字体/行距；**模糊**：结构上像代码块（行 OE + 等宽字体/行距特征） |
| 引用块 | **严格**：单列Table + `hasHeaderRow="false"` + `shadingColor == QuoteBlockConfig.BackgroundColor`（忽略 alpha）且不像代码；**模糊**：单列有底色 Table 且不像代码 |
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

## 嵌入式元数据方案（OneNote `<one:Meta>`）

外部缓存（文件）最大的问题是：**内容与缓存分离**，删除内容后缓存仍然存在，且不利于后续“实时预览/交换”的一致性。
因此，TeXShift 的可逆信息应优先写入 OneNote XML 的 `<one:Meta>`，让“源信息跟着内容走”。

### 存放位置（推荐：Outline 级）

**优先以整个 `Outline` 为“受控单元”**：
- 更像“一个文档双视图（source/render）”，实现简单且直觉一致；
- 一键交换时可原子化地切换该文本框内容；
- 用户手工编辑后可以整体判定过期并丢弃源信息，回落到解析兜底。

可选扩展：对任意选区（若干 `OE`）做细粒度受控，但需要维护边界/锚点，复杂度显著更高，建议后续再做。

### 元数据键命名（建议使用 `texshift-*` 前缀）

不仿照其它插件的前缀（如 `omk-*`），TeXShift 使用自己的键空间以避免冲突：

**受控 Outline 元信息**
- `texshift-schema`: 元数据 schema 版本（例如 `1`）
- `texshift-mode`: `render` / `source`（该 Outline 当前展示形态）
- `texshift-sourceEncoding`: `plain-v1` / `gzip-base64-v1` 等（用于解释 `texshift-source-*`）
- `texshift-source-0...N`: Markdown 源码分片（规范化后的语义等价 Markdown）
- `texshift-sigVersion`: 签名算法版本（例如 `1`）
- `texshift-sig`: 当前内容签名（用于判定是否被用户/OneNote 修改导致源信息过期）

**按需的元素级附加信息（可选）**
- 若后续发现“整段 Markdown”不适合存放（体积/同步等），可退而在特定 `OE` 上仅存不可逆信息：
  - `texshift-math-latex` / `texshift-math-sig`
  - `texshift-mermaid-code` / `texshift-mermaid-sig`
  - `texshift-code-lang` / `texshift-code-sig`

### 版本与兼容性

- `texshift-schema` 是解析入口：未知/不支持版本应直接忽略 Meta，回落到解析兜底（安全默认）。
- `texshift-sourceEncoding` 控制如何从 `texshift-source-*` 还原文本；编码升级需递增版本后缀（例如 `plain-v2`）。
- `texshift-sigVersion` 控制签名算法：未知版本一律视为不匹配（触发“丢弃源信息”）。

### 大文本存储（分片与编码）

`<one:Meta content="...">` 是 attribute，不适合直接保存原始换行/大段文本。
建议策略：
1. **分片**：将源码按固定长度切成 `texshift-source-0...N`，避免单个 attribute 过长（OneNote/Interop 的上限并不透明）。
2. **编码**：
   - `plain-v1`：用 `\n` 表示换行，必要时转义 `\` 与引号（可读性好，便于调试）。
   - `gzip-base64-v1`（可选）：当源码过大时压缩降低 XML 体积，但实现复杂度更高，可后置。

### 过期判定（签名 + 丢弃）

目标：**绝不输出“旧源码”**。一旦不能确定一致，就判定过期并丢弃源信息。

- 正向写入 `render` 时：计算并写入 `texshift-sig`（并记录 `texshift-sigVersion`）。
- 反向/交换时：重新计算当前 `sigNow`。
  - `sigNow == texshift-sig`：认为未被修改 → 可信任 `texshift-source-*`
  - `sigNow != texshift-sig`：认为已被修改（或 OneNote 重写导致不一致）→ **丢弃全部 `texshift-source-*`**，回落到解析兜底，并输出 warning + 占位符（对不可逆元素）。

为降低 OneNote “首次写入后重写 XML” 造成的误判，可在“一键交换/应用渲染”这种低频动作后：
1) 写回 `UpdatePageContent`，2) 立即 `GetPageContent` 读回，3) 用 OneNote 最终落盘形态重新计算并更新 `texshift-sig`。

## 代码块反向转换难点

### 问题
1. 语法高亮输出富 HTML，需要"去高亮"
2. 语言信息 `fenced.Info` 未写入 OneNote XML，反向无法恢复

### 解决方案
1. **HtmlStripper**: 专门的去高亮工具，剥离所有 `<span style='color:...'>` 保留纯文本
2. **语言信息**：受控 Outline 模式下语言随 `texshift-source-*` 一起保留；解析兜底时语言可省略或用元素级 Meta（如 `texshift-code-lang`）补充

## 列表树结构处理

### 问题
正向把列表挂到"前一个容器块"的 OEChildren 下（`MarkdownToOneNoteConverter.cs:171`），
反向若按 XML 层级直接遍历，输出顺序/缩进会错乱。

### 解决方案
**反嵌套/反扁平化策略**：
1. 识别 `OEChildren` 中的列表项
2. 根据 `<Indent>` 配置计算缩进级别
3. 重建 Markdown 列表结构

## 数学公式反向还原

### 现实约束（OneNote 行为）

实测：OneNote 会在回写/读回过程中移除 MathML 的 `<semantics>` / `<annotation encoding="application/x-tex">`，并可能重写甚至截断 MathML 结构，因此无法依赖 annotation 或 MathML→LaTeX 解析来恢复原式。

### 当前策略
1. **正向**：写入 OneNote 前剥离 `<semantics>/<annotation>`（减少 OneNote 重写风险），并将原始 LaTeX 写入 `<one:Meta>`（作为可逆源信息的一部分或元素级附加信息）。
2. **反向**：仅在签名一致时使用 Meta 中的 LaTeX；签名不一致则判定过期，**不输出旧 LaTeX**，改为 warning + 占位符（再尝试解析兜底）。

## 实现阶段

### 阶段1：核心框架
- [ ] `IOneNoteToMarkdownConverter` 接口
- [ ] `IElementHandler` 接口
- [ ] `IOneNoteConverterContext` 接口（注入 `OneNoteStyleConfig` 用于配置值识别）
- [ ] `OneNoteToMarkdownConverter` 骨架
- [ ] `InlineParser` 基础实现（粗体、斜体、删除线、链接）
- [ ] `HtmlStripper` 去高亮工具
- [ ] `TeXShift.Tests.E2E` 增加 `reverse-xml` / `reverse-inplace`（用于阶段性验证与写回验证）

**验证**:
- `InlineParser` 单元测试（粗体、斜体、删除线、链接）
- `HtmlStripper` 单元测试（去除语法高亮 HTML）
- `reverse-xml` E2E 骨架验证

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

### 阶段4：嵌入式元数据（OneNote Meta）
- [ ] 定义 `texshift-*` Meta schema（含版本、编码、分片、签名）
- [ ] 正向转换：在生成 `Outline` 后写入 `texshift-source-*` + `texshift-sig`（支持后续“交换/预览”）
- [ ] 反向转换：Meta 命中且签名一致 → 直接输出 source；否则丢弃并解析兜底
- [ ] 一键交换（写回后读回）：用 OneNote 最终落盘内容更新签名（减少误判）

**验证**：
- 受控 Outline 往返测试
- 签名过期行为、占位符输出
- 正向转换回归测试（防止改坏 `MarkdownToOneNoteConverter`）
- 跨复制/粘贴行为（手动验收）

### 阶段5：特殊元素处理器
- [ ] `MathElementHandler`（Meta 优先；过期/缺失 → warning + 占位符）
- [ ] `MermaidElementHandler`（Meta 优先；过期/缺失 → warning + 占位符）
- [ ] `ImageElementHandler`

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
| `OneNoteMeta/TeXShiftMeta*.cs`（待定） | 读写 `<one:Meta>` + 分片/签名/版本 |

### 需要修改
| 文件 | 修改内容 |
|-----|---------|
| `Services/ServiceContainer.cs` | 添加反向转换器工厂 |
| `Markdown/MarkdownToOneNoteConverter.cs` | 生成 `Outline` 后写入 `texshift-*` Meta（source/签名/版本/分片） |
| `Math/OneNoteMathMLAdapter.cs` | 剥离 `<semantics>/<annotation>`（减少 OneNote 不稳定行为） |
| `Services/ConversionOrchestrator.cs` | 添加反向转换选项 |
| `tests/TeXShift.Tests.E2E/Program.cs` | 增加 `reverse-xml` 子命令 |

## 不可逆元素处理（用户确认）

对于无法还原的元素（源信息缺失/过期、无法解析的 Math/图等）：
- **输出占位符**：避免静默丢失内容
- **收集警告**：在 `ReverseConversionResult.Warnings` 中记录
- **结果呈现**：转换完成后向用户显示警告列表

占位符建议格式（示例）：
- `【TeXShift：公式已被修改，无法还原】`
- `【TeXShift：Mermaid 源码缺失，无法还原】`

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
2. 备用策略失败 → 输出占位符并记录 warning
3. 永不抛异常中断整个转换

## 测试策略

### 单元测试
- `InlineParser` 各种 HTML 样式解析
- `HtmlStripper` 去高亮测试
- 各 `ElementHandler` 独立测试
- `<one:Meta>` 分片/编码/解码
- 签名算法一致性与版本迁移

### 集成测试
- 受控 Outline 的 Meta 读写
- 正向→反向往返保真度

### E2E 测试
- 完整文档往返
- 包含所有元素类型的综合测试
- 无需 Ribbon：E2E 增加 `reverse-xml`（读取 03/04 XML，输出 Markdown）
- 无需 Ribbon：E2E 增加 `reverse-inplace`（基于 OneNote COM：选区 XML → Markdown → 写回 OneNote）






