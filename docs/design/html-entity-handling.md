# HTML 实体处理方案

## 问题
1. OneNote 存储的实体（`&gt;`, `&lt;`）需要被 Markdig 识别为 Markdown 语法
2. LaTeX 公式中的 `&amp;` 需要解码为 `&` 才能被 MathJax 解析
3. 代码块/行内代码中的实体应保持原样（显示字面实体文本）

## 解决方案：三阶段处理

```
Stage 1: Protect     →  Stage 2: Code/Math处理  →  Stage 3: Restore for OneNote HTML
(实体→占位符)           (代码还原,LaTeX解码)        (非代码单层还原,代码双重编码)
```

## 关键代码

### HtmlEntityProcessor.cs
- `Protect()`: 用占位符替换 HTML 实体，返回 entityMap
- `RestoreForCode()`: 静态方法，将占位符还原为原始实体文本（不解码）
- `DecodeForLatex()`: 静态方法，为 Math handlers 解码占位符
- `RestoreForOneNoteHtml()`: 遍历 XML，非代码内容恢复为单层实体，代码内容双重编码

### MarkdownToOneNoteConverter.cs
- `_currentEntityMap`: 存储当前转换的 entityMap
- `DecodeEntityPlaceholders()`: 调用 `HtmlEntityProcessor.DecodeForLatex`
- `RestoreEntityPlaceholders()`: 调用 `HtmlEntityProcessor.RestoreForCode`
- 设置 `_inlineRenderer.EntityDecoder` 委托

### CodeBlockHandler.cs
- 使用 `context.RestoreEntityPlaceholders()` 还原实体文本
- `HtmlEscaper` 在高亮器中进行双重编码（`&lt;` → `&amp;lt;`）

### MathBlockHandler.cs / InlineRenderer.cs
- MathJax 处理前调用 `context.DecodeEntityPlaceholders(latex)`

### 普通富文本的单层实体

OneNote 会把 `<one:T>` 的 CDATA 解释为 HTML。普通文本中的实体应保留一层：

```
用户输入: &lt;br&gt;   (期望显示 "<br>")
CDATA:    &lt;br&gt;
渲染:     <br>          ✓ 字面文本，不会重新解释为标签
```

## 代码块实体的双重编码

OneNote 的 `<one:T>` CDATA 内容被解释为 HTML。为了让实体文本字面显示：

```
用户输入: &lt;div&gt;   (期望显示 "&lt;div&gt;")
         ↓
占位符:  \uE100...\uE101
         ↓
还原:    &lt;div&gt;    (RestoreForCode)
         ↓
双重编码: &amp;lt;div&amp;gt;  (HtmlEscaper)
         ↓
CDATA:   &amp;lt;div&amp;gt;
         ↓
渲染:    &lt;div&gt;     ✓ (OneNote HTML解码一层)
```

## 维护注意事项

1. **不要删除占位符保护**：会导致 Markdig 解码实体，破坏语法检测
2. **不要删除 Math 解码**：会导致 LaTeX 中的 `&` 无法被 MathJax 识别
3. **代码块检测**：通过父 OE 元素的 style 属性匹配代码块字体
4. **行内代码检测**：通过 span 的 background-color 和 font-family 匹配
5. **占位符**：使用 Unicode 私有区域字符 (U+E100-U+F8FF)，冲突风险极低
6. **代码块双重编码**：`RestoreForCode` + `HtmlEscaper` 配合实现
