# LaTeX/MathML 技术债

## 概述

LaTeX 到 MathML 转换使用 MathJax 实现，但 OneNote 对 MathML 的支持有限，导致部分公式无法正确显示。

## 已验证的修复

| 问题 | 修复 | 状态 |
|------|------|------|
| `sin(x)` 等函数公式消失 | 移除 `&#x2061;` + `(` 组合 | ✅ 有效 |
| 括号大小不一致 | 添加 `fence="false"` 到括号 | ✅ 有效 |
| 逗号消失 | 同上（阻止 mfenced 转换） | ✅ 有效 |
| OneNote API 阻塞 | 移除 `stretchy="false"` | ✅ 必须（不移除会导致 API 无响应）|
| 页面更新时函数名变斜体 | 预先拆分多字符标识符为单字符 `<mml:mi>sin</mml:mi>` → `<mml:mrow><mml:mi>s</mml:mi>...</mml:mrow>` | ✅ 有效 |
| 矩阵/分数括号不拉伸 | 检测高内容并转换 `<mo>` 为 `<mfenced>` | ✅ 有效 |
| 间距命令无效 | `<mspace>` 转换为 Unicode 空格字符 | ✅ 有效（精度有限）|

## 未解决的问题

### 1. 连续多个 munderover 结构

**现象**：`\sum_{i=1}^{n} \sum_{j=1}^{m} a_{ij}` 等连续多个带上下标的大运算符显示空白

**正常的公式**：
- 单个求和：`\sum_{i=1}^{n} i` ✅
- 单个求积：`\prod_{i=1}^{n} i` ✅
- 带复杂内容的单个求和：`\sum_{k=0}^{\infty} \frac{x^k}{k!}` ✅

**有问题的公式**：
- 双重求和：`\sum_{i=1}^{n} \sum_{j=1}^{m} a_{ij}` ❌

**分析**：
- OneMark 转换同类公式也失败（空白）
- 只有 OneNote 原生手写公式能正常显示
- 原生公式结构与 MathJax 生成的几乎一致

**尝试过的修复（均无效）**：
- 移除/保留 `display="block"`
- `munderover` 转 `msubsup`
- 添加 `stretchy="false"` 到运算符
- HTML 实体转 Unicode 直接字符
- 移除单元素 mrow 包裹
- 移除零宽空格 span 包裹

**结论**：可能是 OneNote API 对连续 munderover 的限制，暂无解决方案

### 2. LaTeX 换行符 `\\`

**现象**：非矩阵环境中的 `\\` 换行符不生效，`a \\ b` 显示为 `ab`

**原因**：
- MathJax 将 `\\` 转换为 `<mspace linebreak="newline"/>`
- OneNote 不支持 `linebreak` 属性

**可能的解决方案**：
- 将 `<mspace linebreak="newline"/>` 转换为 `<mtable>` 结构
- 复杂度高，可能影响公式对齐

**当前状态**：接受限制，建议用户使用多个公式块代替

**注意**：矩阵内的 `\\` 正常工作，因为矩阵使用 `<mtr>` 行元素

### 3. 页面其他公式变化

**现象**：转换某些公式时，页面上已有的其他公式会被重新解析，导致格式变化

**原因**：不符合 OneNote 预期的 MathML 格式会触发整页重解析

**状态**：需要进一步分析哪些结构会触发此问题

## 支持的公式类型

基于测试，以下类型公式应该能正常工作：
- 基本运算：`a + b`, `a - b`, `a \times b`, `a \div b`
- 分数：`\frac{a}{b}`
- 上下标：`x^2`, `x_i`, `x_i^2`
- 根号：`\sqrt{x}`, `\sqrt[n]{x}`
- 向量：`\vec{a}`
- 希腊字母：`\alpha`, `\beta`, `\pi` 等
- 简单函数：`\sin(x)`, `\cos(x)`, `\log(x)` 等
- 极限（无上标）：`\lim_{x \to 0}`
- 简单括号组合：`(a, b, c)`, `[a, b]`
- 单个求和/求积：`\sum_{i=1}^{n} i`, `\prod_{i=1}^{n} i`
- 矩阵：`pmatrix`, `bmatrix`, `vmatrix` 等（括号自动拉伸）
- 自动拉伸括号：`\left( \frac{a}{b} \right)` 等

## 不支持/有问题的公式类型

- 连续多个带上下标的大运算符：`\sum_{i=1}^{n} \sum_{j=1}^{m}`
- 非矩阵环境的换行：`a \\ b`

## 建议

1. 在用户文档中说明支持范围
2. 对于不支持的公式，考虑降级为图片渲染
3. 后续可研究 OneNote 原生公式的精确格式要求

## 相关文件

- `TeXShift.Core/Math/MathService.cs` - MathML 后处理
- `TeXShift.Core/Resources/Math/mathjax-loader.html` - MathJax 配置
- `docs/design/bracket-stretching.md` - 括号拉伸方案
- `docs/design/mspace-unicode.md` - 间距转换方案
