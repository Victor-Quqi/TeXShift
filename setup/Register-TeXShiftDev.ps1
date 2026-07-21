[CmdletBinding()]
param(
    [ValidateSet("Register", "Status", "Unregister")]
    [string]$Action = "Status",

    [ValidateSet("Debug")]
    [string]$Configuration = "Debug",

    [string]$DllPath
)

$ErrorActionPreference = "Stop"

$RepoRoot = Split-Path -Parent $PSScriptRoot
$ProjectPath = Join-Path $RepoRoot "src\TeXShift.AddIn\TeXShift.AddIn.csproj"
$ManagedRegistrationPath = "Software\TeXShift\DevelopmentRegistration"
$OfficeAddinsRoot = "Software\Microsoft\Office\OneNote\Addins"
$OfficeAddinsDataRoot = "Software\Microsoft\Office\OneNote\AddinsData"
$ManagedMarker = "TeXShift.Register-TeXShiftDev"

function Get-TargetPath {
    if ($DllPath) {
        return [IO.Path]::GetFullPath($DllPath)
    }

    $output = & dotnet msbuild $ProjectPath -getProperty:TargetPath `
        -p:Configuration=$Configuration -p:Platform=x64 -nologo
    if ($LASTEXITCODE -ne 0) {
        throw "Unable to resolve the AddIn output path."
    }

    $targetPath = @($output | Where-Object { $_ -match '\.dll$' } | Select-Object -Last 1)
    if ($targetPath.Count -ne 1) {
        throw "Unable to resolve the AddIn output path."
    }

    return [IO.Path]::GetFullPath($targetPath[0].Trim())
}

function Get-AssemblyIdentity {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Build output not found: $Path"
    }

    $directory = Split-Path -Parent $Path
    $resolver = [ResolveEventHandler] {
        param($sender, $args)

        $name = ([Reflection.AssemblyName]$args.Name).Name + ".dll"
        $candidate = Join-Path $directory $name
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            return [Reflection.Assembly]::LoadFrom($candidate)
        }
        return $null
    }

    [AppDomain]::CurrentDomain.add_AssemblyResolve($resolver)
    try {
        $assembly = [Reflection.Assembly]::LoadFrom($Path)
        $type = $assembly.GetType("TeXShift.AddIn.Connect", $true)
        $guid = $type.GetCustomAttributesData() |
            Where-Object { $_.AttributeType.FullName -eq "System.Runtime.InteropServices.GuidAttribute" } |
            Select-Object -First 1
        $progId = $type.GetCustomAttributesData() |
            Where-Object { $_.AttributeType.FullName -eq "System.Runtime.InteropServices.ProgIdAttribute" } |
            Select-Object -First 1

        if (-not $guid -or -not $progId) {
            throw "COM identity attributes are missing from TeXShift.AddIn.Connect."
        }

        $peKind = [Reflection.PortableExecutableKinds]::NotAPortableExecutableImage
        $machine = [Reflection.ImageFileMachine]::I386
        $assembly.ManifestModule.GetPEKind([ref]$peKind, [ref]$machine)
        if ($machine -ne [Reflection.ImageFileMachine]::AMD64) {
            throw "The AddIn must be built for x64. Detected: $machine"
        }

        return [pscustomobject]@{
            Clsid = "{" + ([Guid]$guid.ConstructorArguments[0].Value).ToString().ToUpperInvariant() + "}"
            ProgId = [string]$progId.ConstructorArguments[0].Value
            Class = $type.FullName
            Assembly = $assembly.GetName().FullName
            Version = $assembly.GetName().Version.ToString()
            RuntimeVersion = $assembly.ImageRuntimeVersion
            CodeBase = if ($Path.StartsWith("\\")) {
                "file:" + $Path.Replace("\", "/")
            }
            else {
                "file:///" + $Path.Replace("\", "/")
            }
            Path = $Path
        }
    }
    finally {
        [AppDomain]::CurrentDomain.remove_AssemblyResolve($resolver)
    }
}

function Open-Registry64 {
    param([Microsoft.Win32.RegistryHive]$Hive)

    return [Microsoft.Win32.RegistryKey]::OpenBaseKey(
        $Hive,
        [Microsoft.Win32.RegistryView]::Registry64)
}

function Assert-Administrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = [Security.Principal.WindowsPrincipal]$identity
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw "Development COM registration requires an elevated PowerShell once. Builds remain non-elevated."
    }
}

function Set-RegistryValues {
    param(
        [Microsoft.Win32.RegistryKey]$Root,
        [string]$Path,
        [hashtable]$Values
    )

    $key = $Root.CreateSubKey($Path)
    try {
        foreach ($name in $Values.Keys) {
            $entry = $Values[$name]
            $key.SetValue($name, $entry.Value, $entry.Kind)
        }
    }
    finally {
        $key.Dispose()
    }
}

function New-StringValue {
    param([AllowEmptyString()][string]$Value)
    return @{ Value = $Value; Kind = [Microsoft.Win32.RegistryValueKind]::String }
}

function New-DWordValue {
    param([int]$Value)
    return @{ Value = $Value; Kind = [Microsoft.Win32.RegistryValueKind]::DWord }
}

function Get-RegistryValue {
    param(
        [Microsoft.Win32.RegistryKey]$Root,
        [string]$Path,
        [AllowEmptyString()][string]$Name
    )

    $key = $Root.OpenSubKey($Path, $false)
    if (-not $key) {
        return $null
    }
    try {
        return $key.GetValue($Name, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
    }
    finally {
        $key.Dispose()
    }
}

function Test-OneNoteStopped {
    $process = Get-Process ONENOTE -ErrorAction SilentlyContinue
    if ($process) {
        throw "Close OneNote before changing the development registration."
    }
}

function Register-DevelopmentAddIn {
    param($Identity)

    Assert-Administrator
    Test-OneNoteStopped
    if ($Identity.ProgId -ne "TeXShift.AddIn.Debug.Connect") {
        throw "Refusing to register a non-Debug COM identity: $($Identity.ProgId)"
    }

    $classesRoot = Open-Registry64 ([Microsoft.Win32.RegistryHive]::LocalMachine)
    $userRoot = Open-Registry64 ([Microsoft.Win32.RegistryHive]::CurrentUser)
    try {
        $clsidPath = "Software\Classes\CLSID\$($Identity.Clsid)"
        $inprocPath = "$clsidPath\InprocServer32"
        $versionPath = "$inprocPath\$($Identity.Version)"
        $progIdPath = "Software\Classes\$($Identity.ProgId)"
        $appIdPath = "Software\Classes\AppID\$($Identity.Clsid)"
        $addinPath = "$OfficeAddinsRoot\$($Identity.ProgId)"

        Set-RegistryValues $classesRoot $clsidPath @{
            "" = New-StringValue $Identity.Class
            "AppID" = New-StringValue $Identity.Clsid
        }
        $inprocValues = @{
            "" = New-StringValue "mscoree.dll"
            "ThreadingModel" = New-StringValue "Both"
            "Class" = New-StringValue $Identity.Class
            "Assembly" = New-StringValue $Identity.Assembly
            "RuntimeVersion" = New-StringValue $Identity.RuntimeVersion
            "CodeBase" = New-StringValue $Identity.CodeBase
        }
        Set-RegistryValues $classesRoot $inprocPath $inprocValues
        Set-RegistryValues $classesRoot $versionPath $inprocValues
        Set-RegistryValues $classesRoot "$clsidPath\ProgId" @{ "" = New-StringValue $Identity.ProgId }
        Set-RegistryValues $classesRoot "$clsidPath\Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}" @{}
        Set-RegistryValues $classesRoot $progIdPath @{ "" = New-StringValue $Identity.Class }
        Set-RegistryValues $classesRoot "$progIdPath\CLSID" @{ "" = New-StringValue $Identity.Clsid }
        Set-RegistryValues $classesRoot $appIdPath @{ "DllSurrogate" = New-StringValue "" }
        Set-RegistryValues $userRoot $addinPath @{
            "FriendlyName" = New-StringValue "TeXShift (Dev)"
            "Description" = New-StringValue "TeXShift development build"
            "LoadBehavior" = New-DWordValue 3
            "CommandLineSafe" = New-DWordValue 1
        }
        Set-RegistryValues $userRoot $ManagedRegistrationPath @{
            "ManagedBy" = New-StringValue $ManagedMarker
            "Clsid" = New-StringValue $Identity.Clsid
            "ProgId" = New-StringValue $Identity.ProgId
            "CodeBase" = New-StringValue $Identity.CodeBase
            "Scope" = New-StringValue "Machine"
        }
    }
    finally {
        $classesRoot.Dispose()
        $userRoot.Dispose()
    }

    Write-Host "Registered $($Identity.ProgId)"
    Write-Host "CodeBase: $($Identity.Path)"
}

function Unregister-DevelopmentAddIn {
    Assert-Administrator
    Test-OneNoteStopped
    $classesRoot = Open-Registry64 ([Microsoft.Win32.RegistryHive]::LocalMachine)
    $userRoot = Open-Registry64 ([Microsoft.Win32.RegistryHive]::CurrentUser)
    try {
        $managedBy = Get-RegistryValue $userRoot $ManagedRegistrationPath "ManagedBy"
        if ($managedBy -ne $ManagedMarker) {
            Write-Host "No managed development registration found."
            return
        }

        $clsid = [string](Get-RegistryValue $userRoot $ManagedRegistrationPath "Clsid")
        $progId = [string](Get-RegistryValue $userRoot $ManagedRegistrationPath "ProgId")
        if (-not $clsid -or -not $progId -or $progId -ne "TeXShift.AddIn.Debug.Connect") {
            throw "The development registration marker is invalid."
        }

        @(
            "Software\Classes\CLSID\$clsid",
            "Software\Classes\$progId",
            "Software\Classes\AppID\$clsid"
        ) | ForEach-Object { $classesRoot.DeleteSubKeyTree($_, $false) }
        @(
            "$OfficeAddinsRoot\$progId",
            "$OfficeAddinsDataRoot\$progId",
            $ManagedRegistrationPath
        ) | ForEach-Object { $userRoot.DeleteSubKeyTree($_, $false) }
    }
    finally {
        $classesRoot.Dispose()
        $userRoot.Dispose()
    }

    Write-Host "Unregistered $progId"
}

function Get-RegistrationStatus {
    param($Identity)

    $checks = [ordered]@{}
    $classesRoot = Open-Registry64 ([Microsoft.Win32.RegistryHive]::LocalMachine)
    $userRoot = Open-Registry64 ([Microsoft.Win32.RegistryHive]::CurrentUser)
    try {
        $clsidPath = "Software\Classes\CLSID\$($Identity.Clsid)"
        $inprocPath = "$clsidPath\InprocServer32"
        $progIdPath = "Software\Classes\$($Identity.ProgId)"
        $appIdPath = "Software\Classes\AppID\$($Identity.Clsid)"
        $addinPath = "$OfficeAddinsRoot\$($Identity.ProgId)"

        $checks["DLL"] = Test-Path -LiteralPath $Identity.Path -PathType Leaf
        $checks["CLSID"] = (Get-RegistryValue $classesRoot "$progIdPath\CLSID" "") -eq $Identity.Clsid
        $checks["AppID"] = (Get-RegistryValue $classesRoot $clsidPath "AppID") -eq $Identity.Clsid
        $checks["Class"] = (Get-RegistryValue $classesRoot $inprocPath "Class") -eq $Identity.Class
        $checks["Assembly"] = (Get-RegistryValue $classesRoot $inprocPath "Assembly") -eq $Identity.Assembly
        $checks["RuntimeVersion"] = (Get-RegistryValue $classesRoot $inprocPath "RuntimeVersion") -eq $Identity.RuntimeVersion
        $registeredCodeBase = [string](Get-RegistryValue $classesRoot $inprocPath "CodeBase")
        $checks["CodeBase"] = $registeredCodeBase -eq $Identity.CodeBase
        $checks["VersionKey"] = (Get-RegistryValue $classesRoot "$inprocPath\$($Identity.Version)" "Assembly") -eq $Identity.Assembly
        $checks["DllSurrogate"] = (Get-RegistryValue $classesRoot $appIdPath "DllSurrogate") -eq ""
        $checks["LoadBehavior"] = (Get-RegistryValue $userRoot $addinPath "LoadBehavior") -eq 3
        $checks["ManagedMarker"] = (Get-RegistryValue $userRoot $ManagedRegistrationPath "ManagedBy") -eq $ManagedMarker
    }
    finally {
        $classesRoot.Dispose()
        $userRoot.Dispose()
    }

    foreach ($check in $checks.GetEnumerator()) {
        $state = if ($check.Value) { "OK" } else { "MISSING" }
        Write-Host ("{0,-16} {1}" -f $check.Key, $state)
    }

    $oneNote = Get-Process ONENOTE -ErrorAction SilentlyContinue
    Write-Host ("{0,-16} {1}" -f "OneNote", $(if ($oneNote) { "running" } else { "stopped" }))

    $surrogate = Get-CimInstance Win32_Process -Filter "Name='dllhost.exe'" -ErrorAction SilentlyContinue |
        Where-Object { $_.CommandLine -and $_.CommandLine.IndexOf($Identity.Clsid, [StringComparison]::OrdinalIgnoreCase) -ge 0 }
    Write-Host ("{0,-16} {1}" -f "COM surrogate", $(if ($surrogate) { "running" } else { "stopped" }))
    Write-Host ("{0,-16} {1}" -f "CodeBase path", $Identity.Path)

    if ($checks.Values -contains $false) {
        return 1
    }
    return 0
}

$targetPath = Get-TargetPath
$identity = Get-AssemblyIdentity $targetPath

switch ($Action) {
    "Register" { Register-DevelopmentAddIn $identity }
    "Unregister" { Unregister-DevelopmentAddIn }
    "Status" { exit (Get-RegistrationStatus $identity) }
}
