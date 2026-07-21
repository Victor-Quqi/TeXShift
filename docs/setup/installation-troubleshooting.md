# TeXShift OneNote Add-in 安装问题排查记录

> 当前入口见 `build.ps1` 与 `setup/Register-TeXShiftDev.ps1`。本文保留 OneNote Click-to-Run 与 `DllSurrogate` 问题的历史证据。

## 排查原则

- 当时的目标是确定加载失败的根本原因，保证用户用旧 MSI 安装包安装完即可正常使用。
- 测试过程中**尽量不要更改注册表**，避免引入干扰变量。如确需修改，改完后必须**立即还原**。

## 环境信息

- OneNote 版本：Microsoft 365 (16.0.19530.20184) 64位
- Office 安装类型：Click-to-Run (C2R)
- 本机和 VM 的 OneNote 版本、Office 类型完全相同

## 症状

- 加载项在 COM 加载项列表中显示，但默认未勾选
- 勾选后点确定 → 变回未勾选（加载失败），无任何错误提示
- `addin-debug.log` 从未生成（静态构造函数从未执行）
- Windows Event Viewer 无相关错误，Office Resiliency 无禁用记录

## 已确认的事实

### 注册表完全正确

- `HKCU\Software\Microsoft\Office\OneNote\Addins\TeXShift.AddIn.Connect`：FriendlyName、Description、LoadBehavior=3、CommandLineSafe=1 均正确
- `HKLM\Software\Classes\CLSID\{1EE8F914-ECBD-4709-92C0-E770851C4966}`：InprocServer32 指向 mscoree.dll，Assembly/Class/RuntimeVersion/CodeBase/ThreadingModel 均正确
- ProgId 双向映射存在，Implemented Categories 存在

### PowerShell COM 激活成功（管理员和非管理员 token 均可）

```powershell
[Activator]::CreateInstance([Type]::GetTypeFromCLSID([Guid]"{1EE8F914-ECBD-4709-92C0-E770851C4966}"))
# 返回: TeXShift.AddIn.Connect
```

### OneNote 的禁用行为（ProcMon 证据）

OneNote 每次启动都会主动禁用加载项，且**禁用发生在进入托管代码之前**：

1. 读取 `LoadBehavior=3`
2. CrashPersistence 写入 `{"AddinExecution":true}`（Length=23）
3. CrashPersistence 写入 `{"AddinExecution":false}`（Length=24）— 两步之间仅约 0.01 秒
4. 写回 `LoadBehavior=2`（禁用）
5. **之后**才去 CreateFile/CreateFileMapping 访问 TeXShift.AddIn.dll

调用栈（ProcMon Stack）确认写回 `LoadBehavior=2` 由 Office 组件发起：
- `ONENOTE.EXE` → `MSO.DLL` → `KernelBase.dll!RegSetValueExW`
- 中间包含 `AppvIsvSubsystems64.dll`（C2R/App-V 虚拟化层）

### COM 注册查找顺序

OneNote 查找 ProgId/CLSID 的顺序：
1. `HKCU\Software\Classes` — 未命中
2. `HKLM\...\ClickToRun\Registry\Machine\...` — 未命中（C2R 虚拟注册表中无对应条目）
3. `HKCR`（合并视图）— 命中

OneNote 读取 CodeBase 后会尝试将路径映射到 Office C2R 的 VFS 路径：
- 探测 `C:\Program Files\Microsoft Office\root\vfs\ProgramFilesX64\VictorQuqi\TeXShift\TeXShift.AddIn.dll` → `PATH NOT FOUND`（失败样本中出现 14 次）
- 回退到真实路径 `C:\Program Files\VictorQuqi\TeXShift\TeXShift.AddIn.dll` → `SUCCESS`

OneNote 在访问 `HKCR\CLSID\{...}\InprocServer32` 时，多次尝试以 **Read/Write** 方式打开，得到 `ACCESS DENIED`（随后以只读方式重试成功）。

### 成功 vs 失败的 ProcMon A/B 对照

| 项目 | 失败（VM） | 成功（本机） |
|------|-----------|------------|
| AddinExecution true→false 间隔 | ~0.01 秒 | ~0.96 秒 |
| 写回 LoadBehavior=2 | 有 | 无 |
| VFS 路径探测 PATH NOT FOUND | 14 次 | 0 次 |
| CodeBase | `file:///C:\Program Files\VictorQuqi\TeXShift\TeXShift.AddIn.dll` | `file:///<repo-root>/src/TeXShift.AddIn/bin/x64/Debug/TeXShift.AddIn.dll` |

### 排除的假设

以下实验均在 VM 上执行，均**未能**解决问题（OneNote 仍然禁用加载项）：

1. **OverrideDefaultDisable=1** — 增加此注册表值后仍被禁用
2. **HKCU COM 镜像** — 将 COM 注册从 HKLM 镜像到 `HKCU\Software\Classes`，仍被禁用
3. **补齐 VFS 路径** — 将 DLL 复制到 `C:\Program Files\Microsoft Office\root\vfs\ProgramFilesX64\VictorQuqi\TeXShift\`，仍被禁用
4. **安装目录 ACL** — 权限正常（Users=RX）
5. **MOTW** — DLL 无 Zone.Identifier 标记
6. **安全策略** — 无 AppLocker/CodeIntegrity/Defender 相关阻止记录，无 Office 禁用加载项策略
7. **进程缓解措施** — BlockDynamicCode、MicrosoftSignedOnly 等均未启用
8. **安装到 Program Files 之外** — 将 DLL 复制到 `C:\TeXShift\`，CodeBase 指向新路径，同时清除 AddinClassifier 缓存，仍被禁用（排除 C2R VFS 路径映射干扰假设）

### AddinClassifier

OneNote 启动过程中会读取 `HKCU\Software\Microsoft\Office\Common\AddinClassifier`，存在与本问题相关的条目（如 `2b1ecbb369856bca0d4ea249c97ba553`），格式为 `v2:0;0;<unix_epoch>`。具体机制未知。

## WinDbg 证据

### 非管理员（可信）

非管理员 WinDbg 通过 Launch Executable 启动 OneNote（进程无特权提升），断点 `mso!CAddInX::HrInternalSetConnect` 返回：

- `eax=0x80040154`（`REGDB_E_CLASSNOTREG` — Class not registered）

**结论：OneNote 的 C2R/App-V 进程内 `CoCreateInstance` 找不到 COM 类注册。** 同一台机器、同一用户的 PowerShell 能成功激活同一 CLSID，说明真实注册表中的 COM 注册正确，但 OneNote 的 App-V 虚拟化层（`AppvIsvSubsystems64.dll`）有独立的 COM 目录，不包含外部 regasm 写入的注册信息。

### ~~管理员~~（不可信）

> **警告：以下证据在管理员权限下获取，OneNote 也以管理员权限运行，COM 注册表视图与正常用户不同，不可信。**

~~管理员 WinDbg 同一断点同样返回 `eax=0x80040154`，但因管理员权限下 COM 查找路径不同，此结果不能独立作为证据。现已被非管理员测试证实。~~

## 本机 vs VM 对照

| 项目 | 本机 | VM |
|------|------|-----|
| Office 版本 | 16.0.19530.20184 | 16.0.19530.20184 |
| 安装类型 | C2R | C2R |
| 注册方式 | VS 构建 | MSI + regasm |
| DLL 路径 | <repo-root>\src\...\Debug\ | C:\Program Files\VictorQuqi\TeXShift\ |
| PowerShell COM | 成功 | 成功 |
| OneNote 加载 | 成功 | 失败 |

## 根本原因（已确认）

**VM 的 MSI 安装缺少 `AppID\DllSurrogate` 注册表条目。**

本机（成功）存在：

```
HKCR\AppID\{1EE8F914-ECBD-4709-92C0-E770851C4966}
    DllSurrogate    REG_SZ    (空字符串)
```

VM（失败）缺少此条目。在 VM 上手动添加后，OneNote 成功加载加载项。

### 机制

- `DllSurrogate=""` 指示 COM 将 DLL 激活到 dllhost.exe（COM Surrogate 进程），而非加载到调用进程内。
- dllhost.exe 不在 C2R/App-V 虚拟化环境内，能正常访问真实 `HKLM\Software\Classes` 注册表，`CoCreateInstance` 成功。
- 缺少此条目时，OneNote 尝试进程内（InprocServer32）加载 → App-V 虚拟化层拦截 COM 激活 → 在虚拟 COM 目录中找不到 CLSID → `REGDB_E_CLASSNOTREG`（0x80040154）→ OneNote 写回 `LoadBehavior=2`（禁用）。
- 本机任务管理器确认 TeXShift 运行在 `DllHost.exe /Processid:{1EE8F914-...}` 中。
- 另一个成功的 OneNote .NET 加载项（[OneMore](https://github.com/stevencohn/OneMore)）也使用了相同的 `DllSurrogate=""` 注册方式。

### 修复

旧 MSI 安装包当时需要额外创建：

```
[HKLM\Software\Classes\AppID\{1EE8F914-ECBD-4709-92C0-E770851C4966}]
"DllSurrogate"=""
```

### 已验证的其他问题

加载成功后发现语言无法切换为中文等其他问题，待后续排查。

## 排查过程中解决的子问题

1. **PowerShell 能激活 COM 但 OneNote 不能** — OneNote C2R 进程的 App-V 虚拟化层拦截进程内 COM 激活，虚拟 COM 目录中无外部 regasm 注册
2. **OneNote 为什么把 AddinExecution 置为 false** — `CoCreateInstance` 失败 → OneNote 认为加载项异常 → 写回 `LoadBehavior=2`
3. **本机能工作但 VM 不能** — 本机有 `AppID\DllSurrogate` → COM 激活走 dllhost.exe（App-V 外）；VM 缺少此条目 → 进程内激活被 App-V 拦截
