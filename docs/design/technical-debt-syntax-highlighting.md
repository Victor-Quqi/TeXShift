# 技术债：语法高亮

## 现状

使用 ColorCode-Universal，存在以下限制：

### 不支持的语言
- Go, Rust, Kotlin, Swift, Ruby

### 部分支持（有缺陷）
| 语言 | 问题 |
|------|------|
| JavaScript | 缺 ES6 (`let`)，无数字/函数高亮 |
| Python | `.` 错误着色，无装饰器支持 |
| HTML | 属性名未识别 |
| JSON | 布尔值未着色 |

### 根本原因
ColorCode 的 Scope 类型有限，语法定义过时。

## 备选方案

| 方案 | 优点 | 缺点 |
|------|------|------|
| TextMateSharp | VSCode 级质量，40+ 语言 | 原生依赖 (Onigwrap)，需管理语法文件 |
| 自定义 Regex | 无依赖，可控 | 开发量大 |

## 建议

1. **短期**: 保持 ColorCode，满足 MVP
2. **长期**: 迁移 TextMateSharp 或等待 ColorCode 更新
