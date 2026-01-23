# Images

目标：覆盖图片在不同块中的位置（段落/列表/引用/表格/内联）以及失效路径的处理。

## Standalone image
![local](<repo-root>/misc/T&MtoN_512.png)

## Inline image (may be degraded to link)
这是一段包含内联图片 ![inline](<repo-root>/misc/T&MtoN_512.png) 的文字。

## Images in lists (may be degraded)
- list item with image: ![local](<repo-root>/misc/T&MtoN_512.png)
1. ordered list with image: ![local](<repo-root>/misc/T&MtoN_512.png)

## Images in blockquotes
> quote line 1
> ![local](<repo-root>/misc/T&MtoN_512.png)
> quote line 3

## Images in tables
| Text | Image |
|:-----|:------|
| left | ![local](<repo-root>/misc/T&MtoN_512.png) |

## Online images
![online](https://avatars.githubusercontent.com/u/80116305?s=64&v=4)

## Invalid paths
![](<repo-root>/misc/does_not_exist.png)
![invalid](invalid_path)
