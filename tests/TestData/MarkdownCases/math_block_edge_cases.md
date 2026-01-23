# Block math edge cases

目标：覆盖相对容易触发解析/渲染问题的块级数学写法（嵌套、符号、cases、left/right 等）。

## Nested structures
$$ \sqrt{\sqrt{\sqrt{x}}} $$
$$ \frac{\frac{\frac{1}{2}}{3}}{4} $$
$$ x^{2^{3^{4}}} $$
$$ x_{i_{j_{k}}} $$

## Special symbols / operators
$$ \infty,\partial,\nabla,\hbar,\aleph $$
$$ \bigcup_{i=1}^{n} A_i \quad \bigcap_{i=1}^{n} A_i $$
$$ \Re(z),\Im(z) $$

## Text mode / cases
$$ x = 1 \text{ if } x > 0 $$
$$ f(x) = \begin{cases} 1 & \text{if } x > 0 \\ 0 & \text{otherwise} \end{cases} $$

## Delimiters / sizing
$$ \left( \frac{a}{b} \right) $$
$$ \left. \frac{df}{dx} \right|_{x=0} $$
