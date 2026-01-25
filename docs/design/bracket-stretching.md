# 括号自动拉伸方案

## 问题

MathJax 生成的 MathML 中，矩阵和 `\left...\right` 括号使用 `<mo>` 元素，OneNote 不会自动拉伸这些括号以匹配内容高度。

**MathJax 输出**:
```xml
<mml:mrow>
  <mml:mo>(</mml:mo>
  <mml:mfrac>...</mml:mfrac>
  <mml:mo>)</mml:mo>
</mml:mrow>
```

**OneNote 需要**:
```xml
<mml:mfenced>
  <mml:mrow>
    <mml:mfrac>...</mml:mfrac>
  </mml:mrow>
</mml:mfenced>
```

## 解决方案

在 `MathService.ConvertMatrixBracketsToMfenced()` 中，检测 `<mrow><mo>BRACKET</mo>...CONTENT...<mo>BRACKET</mo></mrow>` 模式，当内容包含高元素时转换为 `<mfenced>`。

### 支持的括号类型

| 括号 | LaTeX | 用途 |
|------|-------|------|
| `()` | `\left( \right)`, `pmatrix` | 圆括号 |
| `[]` | `\left[ \right]`, `bmatrix` | 方括号 |
| `{}` | `\left\{ \right\}`, `Bmatrix` | 花括号 |
| `||` | `\left| \right|`, `vmatrix` | 竖线/行列式 |

### 触发转换的高元素

- `mtable` - 矩阵
- `mfrac` - 分数
- `msqrt`, `mroot` - 根号
- `munder`, `mover`, `munderover` - 上下标结构
- `mfenced` - 嵌套括号

### 单边括号（cases 环境）

`cases` 环境产生 `<mrow><mo>{</mo>...<mo stretchy="true"></mo></mrow>`（右边为空的不可见边界），需转换为 `<mfenced open="{" close="">...`。

### 不转换的情况

普通括号如 `(a + b)` 不包含高元素，保持原样并添加 `fence="false"` 防止 OneNote 错误转换。

## 已知限制

同类型嵌套括号 `\left(\left(\frac{}{}\right)\right)` 可能匹配不正确，但此写法很少见。

## 相关文件

- `TeXShift.Core/Math/MathService.cs` - `ConvertMatrixBracketsToMfenced()`
