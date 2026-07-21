# Block math edge cases

测试深层嵌套、特殊符号、`text`/`cases` 与定界符变体。

## 深层嵌套的根式、分数与上下标
$$ \sqrt{\sqrt{\sqrt{x}}} $$
$$ \frac{\frac{\frac{1}{2}}{3}}{4} $$
$$ x^{2^{3^{4}}} $$
$$ x_{i_{j_{k}}} $$

## 特殊符号与大运算符
$$ \infty,\partial,\nabla,\hbar,\aleph $$
$$ \bigcup_{i=1}^{n} A_i \quad \bigcap_{i=1}^{n} A_i $$
$$ \Re(z),\Im(z) $$
$$ \sum_{i=1}^{n} \sum_{j=1}^{m} a_{ij} $$

## `\text` 与 `cases` 环境
$$ x = 1 \text{ if } x > 0 $$
$$ f(x) = \begin{cases} 1 & \text{if } x > 0 \\ 0 & \text{otherwise} \end{cases} $$

## 自动伸缩与单边定界符
$$ \left( \frac{a}{b} \right) $$
$$ \left. \frac{df}{dx} \right|_{x=0} $$
