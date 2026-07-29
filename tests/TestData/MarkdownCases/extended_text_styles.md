# Extended Inline Styles

## Basic styles

Highlight: ==important text==

Underline: ++underlined text++

Superscript and subscript: x^2^, H~2~O

Adjacent styles: ==重点==++下划线++^sup^~sub~

## Typical combinations

==**Highlighted bold text** and [a highlighted link](https://example.com)==

++*Underlined italic text* and ~~obsolete H~2~O~~++

Math keeps precedence: $x_i^2 + y_{total}$ and ^annotation^.

==Highlighted `inline code`== and `==literal== ++literal++ ^literal^ ~literal~`.

==First highlighted line<br>second highlighted line==

## Block contexts

### Heading with ==highlight== and x^2^

- List item with ++underline++ and H~2~O
- List item with ==**nested emphasis**==
- [x] Task with ^superscript^ and ~~old H~2~O~~

> Quote with ==highlight==, ++underline++, and `code`.

| Context | Mixed content |
| --- | --- |
| Link | ==[documentation](https://example.com/docs)== |
| Formula | $E = mc^2$ and x^note^ |
| Decorations | ++*current* and ~~old H~2~O~~++ |

## Inline HTML style tags

<mark data-source="html">**Review [the change](https://example.com/review)** and `inline code`</mark>; <ins>*new H<sub>2</sub>O*</ins>; <del>old x<sup>2</sup></del>; math stays $x_i^2$.

- Aliases: <u>underlined</u>, <s>obsolete</s>, <MARK>highlighted</MARK>.
- Script positions: x<sup>2</sup> + H<sub>2</sub>O.
