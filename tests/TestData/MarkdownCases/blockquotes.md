# Blockquotes

目标：覆盖引用块嵌套、空引用、引用中包含列表/代码块、以及与普通段落紧邻的情况。

## Basic
> 这是一个简单的引用块。

## Nesting (3 levels)
> 第一层引用
> > 第二层引用
> > > 第三层引用（最深层）

## Empty quote
>

## Quote with list inside
> - list item 1
> - list item 2
>   - nested list item

## Quote with fenced code block
> ```python
> def hello():
>     return "Hello from quote"
> ```

## Quote adjacent to paragraphs (no blank lines)
普通段落（上方）
> 紧邻上方段落的引用块
普通段落（下方）

## Quote with image
> ![local](<repo-root>/misc/T&MtoN_512.png)
