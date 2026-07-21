# Block math: whitespace + newlines

测试 `$$` 的紧贴、空格和多行定界符解析形态，以及 `<`/`>` XML 敏感字符的实体处理。

## `$$` 紧贴公式内容
$$\begin{pmatrix} a & b \\ c & d \end{pmatrix}$$

## `$$` 与公式内容之间留空格
$$ \begin{pmatrix} a & b \\ c & d \end{pmatrix} $$

## 独立成行的 `$$` 多行公式块
$$
\begin{pmatrix}
a_{11} & a_{12} \\
a_{21} & a_{22}
\end{pmatrix}
$$

## 单行比较运算符的 XML 实体处理
$$ a < b $$
$$ a > b $$
$$ a \lt b $$
$$ a \gt b $$
$$ a \leq b $$
$$ a \geq b $$
$$ a \neq b $$

## 多行比较运算符的 XML 实体处理
$$
a < b
$$
