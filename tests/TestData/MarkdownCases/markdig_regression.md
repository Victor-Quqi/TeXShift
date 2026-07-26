# Markdig regression

## Inline parsing

Plain **bold**, *italic*, and ~~strikethrough~~ text.

Extended emphasis: ==marked==, ++inserted++, ^superscript^, and ~subscript~.

Single signs: charge^+^ and ion~+~.

HTML supports **bold** and *italic*.

*[HTML]: HyperText Markup Language

中文段落包含 **粗体**、*斜体* 和 `inline<br>code`。

Line A<br>Line B<br />Line C

Inline math keeps delimiters isolated: $x_{total}^2 + y_i$.

## Pipe table boundaries

| Component | Per query | Per 1,000 queries |
| --- | ---: | ---: |
| Embedding | ~$0.00001 | ~$0.01 |
| LLM | ~$0.0015 | ~$1.50 |
| **Total** | **~$0.0015** | **~$1.50** |

## Quote and ordered list

> 3. quoted third item
> 4. quoted fourth item with **bold text**

7. outer seventh item
8. outer eighth item
   - nested bullet
   - nested bullet with `code`
