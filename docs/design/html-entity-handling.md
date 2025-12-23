# HTML 实体处理方案

## 问题
1. OneNote 存储的实体（`&gt;`, `&lt;`）需要被 Markdig 识别为 Markdown 语法
2. LaTeX 公式中的 `&amp;` 需要解码为 `&` 才能被 MathJax 解析
3. 代码块/行内代码中的实体应保持原样

## 解决方案：三阶段解码

```
Stage 1: Protect     →  Stage 2: Math Decode  →  Stage 3: Restore & Decode
(实体→占位符)           (LaTeX 占位符解码)         (非代码解码,代码保留)
```

## 关键代码

### HtmlEntityProcessor.cs
- `Protect()`: 用占位符替换 HTML 实体，返回 entityMap
- `RestoreAndDecode()`: 遍历 XML，非代码内容解码，代码内容保留原实体
- `DecodeForLatex()`: 静态方法，为 Math handlers 解码占位符

### MarkdownToOneNoteConverter.cs
- `_currentEntityMap`: 存储当前转换的 entityMap
- `DecodeLatexEntities()`: 调用 `HtmlEntityProcessor.DecodeForLatex`
- 设置 `_inlineRenderer.EntityDecoder` 委托

### MathBlockHandler.cs / InlineRenderer.cs
- MathJax 处理前调用 `context.DecodeLatexEntities(latex)`

## 维护注意事项

1. **不要删除占位符保护**：会导致 Markdig 解码实体，破坏语法检测
2. **不要删除 Math 解码**：会导致 LaTeX 中的 `&` 无法被 MathJax 识别
3. **代码块检测**：通过 `Cell.shadingColor` 属性判断是否在代码块内
4. **占位符**：使用 Unicode 私有区域字符 (U+E100-U+F8FF)，冲突风险极低
