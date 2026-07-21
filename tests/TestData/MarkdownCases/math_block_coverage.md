# Block math coverage (`$$...$$`)

测试块级公式的结构类型覆盖。

## 基本运算、分数与根式
$$ x + y = z $$
$$ \frac{a+b}{c+d} = 1 $$
$$ \sqrt{x^2 + y^2} $$

## 求和、乘积与极限
$$ \sum_{i=1}^{n} i = \frac{n(n+1)}{2} $$
$$ \prod_{i=1}^{n} i $$
$$ \lim_{x \to 0} \frac{\sin x}{x} = 1 $$

## 积分
$$ \int_0^1 x\,dx = \frac{1}{2} $$
$$ \int_{0}^{\pi} \sin(x)\,dx = 2 $$

## 矩阵
$$ \begin{pmatrix} a & b \\ c & d \end{pmatrix} $$

## 数字开头的公式（字体回归）

同时覆盖首 token 为数字、普通括号不拉伸（`fence=false`）、`\left...\right` 拉伸为 `mfenced`，以及下标。

$$240 + 325 + 2(290) + \frac{4}{45}(20) - \left(4 + \frac{4}{45}\right)T_1 + \frac{6400}{45} = 0$$

## 括号内逗号（防 mfenced 吞逗号）

$$(a, b, c)$$
$$[a, b]$$
