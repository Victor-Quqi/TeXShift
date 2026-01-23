# Task lists (checkboxes)

目标：覆盖 `- [ ]` / `- [x]`、嵌套、与其他块混排、中文与特殊字符。

## Basic
- [ ] 未完成任务
- [x] 已完成任务

## Inline formatting inside task item
- [ ] **bold** *italic* `code` [link](https://example.com)
- [x] ~~strikethrough~~ then text

## Nesting
- [ ] Parent task
  - [ ] Child task A
  - [x] Child task B
  - [ ] Child task C

## Mixed with other list types
- [ ] Task item
- Normal bullet item (not a task)
1. Ordered item 1
2. Ordered item 2
- [x] Task after ordered list

## Mixed with other blocks
> Quote before tasks

- [ ] Task after quote

```csharp
// fenced code block between tasks
Console.WriteLine("Hello from code fence");
```

- [x] Task after code fence

---

## Chinese + special characters
- [ ] 中文：检查“缩进/换行/标点”是否正常（，。！？《》）
- [x] Special chars: <>&"'\ @#$%^*() / \\ | = + - _

