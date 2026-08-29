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
| 公式首 token 为数字时字体错误（继承正文字体而非 Cambria Math） | 移除前置零宽空格哨兵 span，仅保留尾部哨兵 | ✅ 有效 |

## 未解决的问题

### 1. LaTeX 换行符 `\\`

**现象**：非矩阵环境中的 `\\` 换行符不生效，`a \\ b` 显示为 `ab`

**原因**：
- MathJax 将 `\\` 转换为 `<mspace linebreak="newline"/>`
- OneNote 不支持 `linebreak` 属性

**可能的解决方案**：
- 将 `<mspace linebreak="newline"/>` 转换为 `<mtable>` 结构
- 复杂度高，可能影响公式对齐

**当前状态**：接受限制，建议用户使用多个公式块代替

**注意**：矩阵内的 `\\` 正常工作，因为矩阵使用 `<mtr>` 行元素

### 2. 页面其他公式变化

**现象**：转换某些公式时，页面上已有的其他公式会被重新解析，导致格式变化

**原因**：不符合 OneNote 预期的 MathML 格式会触发整页重解析

**已排除**：与"写入动作本身"无关。实测反复以相同载荷回写已转换页面（只回写 Outline / 回写整页两种形态，以及每轮向 Outline 新增一个原生公式），`math_block_coverage.md` 的 12 个公式（含矩阵、积分、`mfenced` 拉伸、长式子）CDATA 逐字节不变，`<br>` 计数恒为 0。新增公式也不会扰动同一 Outline 内已有的公式。

**状态**：需要进一步分析哪些结构会触发此问题。复现思路：喂入含可疑公式的 Markdown，转换后比对页内其他公式的 CDATA 哈希。

## 支持的公式类型

以下公式类型经测试可正常工作：
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
- 连续多个带上下限的大运算符：`\sum_{i=1}^{n} \sum_{j=1}^{m} a_{ij}`
- 矩阵：`pmatrix`, `bmatrix`, `vmatrix` 等（括号自动拉伸）
- 自动拉伸括号：`\left( \frac{a}{b} \right)` 等

## 不支持/有问题的公式类型

- 非矩阵环境的换行：`a \\ b`

## 相关文件

- `TeXShift.Core/Math/MathService.cs` - LaTeX 转 MathML（WebView2 + MathJax）
- `TeXShift.Core/Math/OneNoteMathMLAdapter.cs` - MathML 后处理与 OneNote 适配
- `TeXShift.Core/Resources/Math/mathjax-loader.html` - MathJax 配置
- `docs/design/bracket-stretching.md` - 括号拉伸方案
- `docs/design/mspace-unicode.md` - 间距转换方案
