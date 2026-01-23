# Block math: whitespace + newlines

目标：覆盖 `$$` 与内容“紧贴/空格/换行”、以及矩阵/比较符号在不同换行方式下的表现（曾经常见问题点）。

## Tight `$$` (no spaces)
$$\begin{pmatrix} a & b \\ c & d \end{pmatrix}$$

## Spaced `$$`
$$ \begin{pmatrix} a & b \\ c & d \end{pmatrix} $$

## Multi-line block
$$
\begin{pmatrix}
a_{11} & a_{12} \\
a_{21} & a_{22}
\end{pmatrix}
$$

## Comparisons (single-line)
$$ a < b $$
$$ a > b $$
$$ a \lt b $$
$$ a \gt b $$
$$ a \leq b $$
$$ a \geq b $$
$$ a \neq b $$

## Comparisons (multi-line)
$$
a < b
$$
