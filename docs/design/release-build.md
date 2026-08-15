# 构建与发布

`build.ps1` 是本机、agent 和 CI 的统一入口。Visual Studio 仅作为可选编辑器。

```powershell
.\build.ps1 -Target Build -Configuration Debug
.\build.ps1 -Target Test -Configuration Debug
.\build.ps1 -Target CI -Configuration Release
```

脚本使用 `global.json` 固定 .NET SDK，按 lock file 恢复 NuGet 包，并校验 MathJax 下载哈希。CI 只执行 Release 构建、快速单元测试和 staging 校验；OneNote E2E 留在本机手动执行。

首次开发注册需要关闭 OneNote，并在管理员 PowerShell 中运行：

```powershell
.\setup\Register-TeXShiftDev.ps1 -Action Register
```

注册项指向固定的 Debug 输出路径，后续构建无需提权或重新注册。重启 OneNote 后加载新版。Debug 使用独立 CLSID/ProgID，可与已安装的 Release 并存；此时会出现两套 Ribbon。用管理员 PowerShell 执行 `-Action Unregister` 可移除 Debug 版。

Release 打包需要 Inno Setup 7 和本地强名称密钥。密钥可放在仓库根目录的 `texshift_key.snk`，或通过 `TEXSHIFT_SIGNING_KEY_PATH` 指定；两者均不得提交。

```powershell
$Version = Read-Host "Release version (x.y.z)"
.\build.ps1 -Target Package -Configuration Release -Version $Version
```

输出位于 `artifacts/package/`，包含单文件安装器和 SHA-256 文件。打包会检查 Release COM 身份、x64、强名称、依赖文件、资源清单和 `THIRD-PARTY-NOTICES.md`。安装器写入机器级 COM/OneNote 注册，安装和卸载需要管理员权限。卸载时可选择删除安装用户的默认 TeXShift 数据目录；用户配置的外部调试输出目录保留。第三方许可证声明随安装文件写入应用目录。

发布正文来自本地忽略的 `.release/release-template.local.md`，仓库不配置自动发布任务。
