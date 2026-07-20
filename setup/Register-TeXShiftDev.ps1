#Requires -RunAsAdministrator

[CmdletBinding()]
param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Debug"
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$dllPath = Join-Path $repoRoot "src\TeXShift.AddIn\bin\x64\$Configuration\TeXShift.AddIn.dll"
$regAsmPath = Join-Path $env:WINDIR "Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"

if (-not (Test-Path -LiteralPath $dllPath -PathType Leaf)) {
    throw "Build output not found: $dllPath"
}

if (-not (Test-Path -LiteralPath $regAsmPath -PathType Leaf)) {
    throw "64-bit RegAsm not found: $regAsmPath"
}

& $regAsmPath $dllPath /codebase
if ($LASTEXITCODE -ne 0) {
    throw "RegAsm failed with exit code $LASTEXITCODE."
}

function Ensure-RegistryKey {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path | Out-Null
    }
}

$clsid = "{1EE8F914-ECBD-4709-92C0-E770851C4966}"
$addinKey = "HKCU:\Software\Microsoft\Office\OneNote\Addins\TeXShift.AddIn.Connect"
$clsidKey = "HKLM:\SOFTWARE\Classes\CLSID\$clsid"
$appIdKey = "HKLM:\SOFTWARE\Classes\AppID\$clsid"

# Register the OneNote add-in for the current user.
Ensure-RegistryKey -Path $addinKey
New-ItemProperty -Path $addinKey -Name "FriendlyName" -Value "TeXShift" -PropertyType String -Force | Out-Null
New-ItemProperty -Path $addinKey -Name "Description" -Value "TeXShift Dev Add-in" -PropertyType String -Force | Out-Null
New-ItemProperty -Path $addinKey -Name "LoadBehavior" -Value 3 -PropertyType DWord -Force | Out-Null
New-ItemProperty -Path $addinKey -Name "CommandLineSafe" -Value 1 -PropertyType DWord -Force | Out-Null

# Keep COM activation outside the Office Click-to-Run virtualized process.
Ensure-RegistryKey -Path $clsidKey
New-ItemProperty -Path $clsidKey -Name "AppID" -Value $clsid -PropertyType String -Force | Out-Null
Ensure-RegistryKey -Path $appIdKey
New-ItemProperty -Path $appIdKey -Name "DllSurrogate" -Value "" -PropertyType String -Force | Out-Null

$assemblyName = [Reflection.AssemblyName]::GetAssemblyName($dllPath).FullName
Write-Host "Registered $assemblyName"
Write-Host "CodeBase: $dllPath"