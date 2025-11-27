# Release 构建

## 前置条件

- Visual Studio 2022 + Microsoft Visual Studio Installer Projects 扩展
- .NET Framework 4.8 SDK
- 目标电脑需要：
  - .NET Framework 4.8
  - WebView2 Runtime

## 构建步骤

1. **切换配置**：工具栏选择 `Release | x64`

2. **更新版本号**：
   - 单击 TeXShiftSetup 项目，按 F4
   - 修改 `Version` 属性（如 `1.0.0` → `1.0.1`）
   - 提示更新 ProductCode 时选"是"

3. **生成解决方案**：右键解决方案 → 重新生成解决方案

4. **MSI 输出位置**：`TeXShiftSetup\Release\TeXShift.msi`

## Setup 项目配置说明

Application Folder 包含两个项目输出：

- **Primary Output from TeXShift**：主 DLL 及依赖
- **Content Files from TeXShift**：MathJax 资源文件（通过 csproj 中的 Content Include 自动包含）

## 关键属性（F4 属性窗口）

| 属性 | 说明 |
|------|------|
| ProductName | 安装后显示的程序名 |
| Version | 版本号，每次发布需递增 |
| Manufacturer | 制造商/作者 |
| TargetPlatform | 必须是 x64 |

## 注意事项

- MathJax 文件通过 `TeXShift.csproj` 的 Content 配置自动复制到输出目录
- 修改 MSI 文件名：右键 TeXShiftSetup → 属性 → Build → Output file name
