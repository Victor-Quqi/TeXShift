# Math formulas with curly braces

测试 LaTeX 公式中的花括号（下标/上标）与标题、段落、列表、引用块、表格等 Markdown 结构互操作。

## 标题中的公式
### 热阻 $R''_{total}$
#### 下标 $T_{cu, interface}$
##### 复合下标 $R''_{contact}$

## 段落中的公式
总热阻公式：$R''_{total} = \frac{T_{cu, interface} - T_{\infty}}{\dot{q}}$

铝板导热热阻：$R''_{al} = \frac{L_{al}}{k_{al}}$

连续多个下标：$A_{i_{j_{k}}}$ 和 $B^{x^{y^z}}$

## 列表中的公式
- 接触热阻 $R''_{contact}$
- 对流热阻 $R''_{conv}$
- 传热系数 $h_c = \frac{1}{R''_{contact}}$

## 引用块中的公式
> 根据公式 $\dot{q} = \frac{\Delta T}{R''_{total}}$，
> 我们可以计算出 $R''_{total} = 0.015094\text{ m}^2\cdot\text{K/W}$

## 表格中的公式
| 符号 | 公式 |
|------|------|
| 总热阻 | $R''_{total}$ |
| 铝板热阻 | $R''_{al}$ |
| 对流热阻 | $R''_{conv}$ |
