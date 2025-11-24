# HTML 实体处理方案

## 问题
OneNote 存储的实体（`&gt;`, `&lt;`）需要被 Markdig 识别为 Markdown 语法（如 `>` 引用块），但用户输入的实体（`&lt;` → `<`）不能被二次解码。

## 解决方案
四步机制：HtmlDecode → 占位符保护 → Markdig 解析 → 恢复占位符

## 关键代码

### MarkdownConverter.cs
- `HtmlEntityRegex`：正则表达式，匹配需要保护的 HTML 实体
- `ConvertToOneNoteXml()`：实现四步处理流程
- `ProtectHtmlEntities()`：用占位符（`\uFFFD{n}\uFFFD`）替换实体
- `RestoreHtmlEntities()`：恢复占位符为原始实体

### ContentReader.cs
- `ProcessOE()`：删除了 `HtmlDecode` 调用（`XElement.Value` 已自动解码）

## 维护注意事项

1. **不要删除 HtmlDecode 步骤**：会导致 Markdown 语法（`>`, `#`, `-` 等）无法识别
2. **不要删除占位符机制**：会导致用户输入的 HTML 实体被 Markdig 二次解码丢失
3. **占位符冲突风险极低**：使用 Unicode 替换字符 (U+FFFD)，正常文本中罕见
4. **修改支持的实体**：修改 `HtmlEntityRegex` 的正则表达式
