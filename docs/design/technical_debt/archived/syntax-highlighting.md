# 技术债：语法高亮

## 已解决 (2026-01)

已从 ColorCode 迁移到 TextMateSharp 2.0.2。

### 改进

- 40+ 语言支持 (包括 Go, Rust, Kotlin, Swift, Ruby)
- VS Code 级别语法精度
- 内置 Dark Plus 主题
- 无需手动维护语言别名映射

### 依赖

- TextMateSharp 2.0.2
- TextMateSharp.Grammars 2.0.2

---

## 历史记录

### 原问题 (ColorCode)

使用 ColorCode-Universal 存在以下限制：

**不支持的语言**: Go, Rust, Kotlin, Swift, Ruby

**部分支持（有缺陷）**:

| 语言 | 问题 |
|------|------|
| JavaScript | 缺 ES6 (`let`)，无数字/函数高亮 |
| Python | `.` 错误着色，无装饰器支持 |
| HTML | 属性名未识别 |
| JSON | 布尔值未着色 |

**根本原因**: ColorCode 的 Scope 类型有限，语法定义过时。
