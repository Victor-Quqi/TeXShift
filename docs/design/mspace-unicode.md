# MathML 间距转 Unicode 方案

## 问题

OneNote 不支持 MathML 的 `<mspace width="...">` 元素，LaTeX 间距命令 (`\,`, `\:`, `\;`, `\quad`, `\qquad`) 无法正确显示。

## 解决方案

在 `MathService.ConvertMspaceToUnicode()` 中，将 `<mspace>` 元素转换为 Unicode 空格字符。

### 映射关系

| LaTeX | em 值 | Unicode 字符 |
|-------|-------|-------------|
| `\qquad` | ≥1.5em | `\u2003\u2003` (两个 em space) |
| `\quad` | ≥0.8em | `\u2003` (em space) |
| `\;` | ≥0.25em | `\u2004` (three-per-em space) |
| `\:` | ≥0.2em | `\u205F` (medium mathematical space) |
| `\,` | ≥0.1em | `\u2009` (thin space) |

### 限制

- 负间距 (`\!`) 无法模拟，直接忽略
- 间距精度不如原生 MathML

## 相关文件

- `TeXShift.Core/Math/MathService.cs` - `ConvertMspaceToUnicode()`, `GetUnicodeSpaceForWidth()`
