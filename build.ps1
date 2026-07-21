[CmdletBinding()]
param(
    [ValidateSet("Clean", "Restore", "Build", "Test", "Stage", "Verify", "Package", "E2E", "CI", "All")]
    [string]$Target = "Build",

    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Debug",

    [ValidatePattern("^\d+\.\d+\.\d+$")]
    [string]$Version,

    [string]$SigningKeyPath,

    [string]$InnoSetupCompiler,

    [ValidateSet("convert", "reverse-xml")]
    [string]$E2ECommand = "convert",

    [string]$E2EInput,

    [string]$E2EOutput
)

$ErrorActionPreference = "Stop"

# A PSModulePath inherited from a PowerShell 7 host shadows Windows PowerShell's
# built-in modules and breaks cmdlet autoloading (e.g. Get-FileHash); make sure
# $PSHOME\Modules is searched first.
if ($PSVersionTable.PSEdition -ne "Core") {
    $env:PSModulePath = (Join-Path $PSHome "Modules") + ";" + $env:PSModulePath
}

$RepoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$AddInProject = Join-Path $RepoRoot "src\TeXShift.AddIn\TeXShift.AddIn.csproj"
$E2EProject = Join-Path $RepoRoot "tests\TeXShift.Tests.E2E\TeXShift.Tests.E2E.csproj"
$UnitTestProject = Join-Path $RepoRoot "tests\TeXShift.Core.Tests\TeXShift.Core.Tests.csproj"
$ArtifactsRoot = Join-Path $RepoRoot "artifacts"
$StageRoot = Join-Path $ArtifactsRoot "bin"
$TestResultsRoot = Join-Path $ArtifactsRoot "test-results"
$DiagnosticsRoot = Join-Path $ArtifactsRoot "diagnostics"
$PackageRoot = Join-Path $ArtifactsRoot "package"
$MathJaxSetup = Join-Path $RepoRoot "setup-mathjax.ps1"
$InstallerScript = Join-Path $RepoRoot "setup\TeXShift.iss"

function Invoke-Dotnet {
    param([string[]]$Arguments)

    Write-Host ("dotnet " + ($Arguments -join " "))
    & dotnet @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "dotnet failed with exit code $LASTEXITCODE."
    }
}

function Invoke-PowerShellFile {
    param(
        [string]$Path,
        [string[]]$Arguments = @()
    )

    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $Path @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "$Path failed with exit code $LASTEXITCODE."
    }
}

function Get-BuildProperties {
    $properties = @("-p:Platform=x64")
    if ($Version) {
        $properties += "-p:Version=$Version"
        $properties += "-p:FileVersion=$Version.0"
        $properties += "-p:InformationalVersion=$Version"
    }

    $resolvedSigningKey = $SigningKeyPath
    if (-not $resolvedSigningKey) {
        $resolvedSigningKey = $env:TEXSHIFT_SIGNING_KEY_PATH
    }
    if ($resolvedSigningKey) {
        $resolvedSigningKey = [IO.Path]::GetFullPath($resolvedSigningKey)
        $properties += "-p:TeXShiftSigningKeyPath=$resolvedSigningKey"
    }

    return $properties
}

function Invoke-Restore {
    param([switch]$Locked)

    Invoke-PowerShellFile $MathJaxSetup
    $projects = @($AddInProject, $E2EProject)
    if (Test-Path -LiteralPath $UnitTestProject -PathType Leaf) {
        $projects += $UnitTestProject
    }

    foreach ($project in $projects) {
        $arguments = @("restore", $project, "-p:Platform=x64")
        if ($Locked) {
            $arguments += "--locked-mode"
        }
        Invoke-Dotnet $arguments
    }
}

function Invoke-Clean {
    $projects = @($AddInProject, $E2EProject)
    if (Test-Path -LiteralPath $UnitTestProject -PathType Leaf) {
        $projects += $UnitTestProject
    }

    foreach ($project in $projects) {
        Invoke-Dotnet @("clean", $project, "-c", $Configuration, "-p:Platform=x64", "--verbosity", "quiet")
    }
}

function Invoke-Build {
    $properties = Get-BuildProperties
    foreach ($project in @($AddInProject, $E2EProject)) {
        Invoke-Dotnet (@("build", $project, "-c", $Configuration, "--no-restore") + $properties)
    }

    if (Test-Path -LiteralPath $UnitTestProject -PathType Leaf) {
        Invoke-Dotnet (@("build", $UnitTestProject, "-c", $Configuration, "--no-restore") + $properties)
    }
}

function Invoke-UnitTests {
    if (-not (Test-Path -LiteralPath $UnitTestProject -PathType Leaf)) {
        throw "Unit test project not found: $UnitTestProject"
    }

    Reset-ArtifactDirectory $TestResultsRoot
    $properties = Get-BuildProperties
    Invoke-Dotnet (@(
        "test", $UnitTestProject,
        "-c", $Configuration,
        "--no-build",
        "--no-restore",
        "--logger", "trx;LogFileName=TeXShift.Core.Tests.trx",
        "--results-directory", $TestResultsRoot
    ) + $properties)
}

function Reset-ArtifactDirectory {
    param([string]$Path)

    $fullArtifactsRoot = [IO.Path]::GetFullPath($ArtifactsRoot).TrimEnd('\') + '\'
    $fullPath = [IO.Path]::GetFullPath($Path)
    if (-not $fullPath.StartsWith($fullArtifactsRoot, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to reset a directory outside artifacts: $fullPath"
    }

    if (Test-Path -LiteralPath $fullPath) {
        Remove-Item -LiteralPath $fullPath -Recurse -Force
    }
    New-Item -ItemType Directory -Path $fullPath -Force | Out-Null
}

function Copy-StageDirectory {
    param(
        [string]$OutputRoot,
        [string]$RelativePath
    )

    $source = Join-Path $OutputRoot $RelativePath
    if (-not (Test-Path -LiteralPath $source -PathType Container)) {
        throw "Required output directory not found: $source"
    }

    $relativeParent = Split-Path -Parent $RelativePath
    $destinationParent = if ($relativeParent) { Join-Path $StageRoot $relativeParent } else { $StageRoot }
    New-Item -ItemType Directory -Path $destinationParent -Force | Out-Null
    Copy-Item -LiteralPath $source -Destination $destinationParent -Recurse -Force
}

function Test-StagedPayload {
    $requiredFiles = @(
        "TeXShift.AddIn.dll",
        "TeXShift.AddIn.dll.config",
        "TeXShift.Core.dll",
        "Office.dll",
        "MaterialDesignThemes.Wpf.dll",
        "Microsoft.Web.WebView2.Core.dll",
        "Microsoft.Web.WebView2.Wpf.dll",
        "TextMateSharp.dll",
        "TextMateSharp.Grammars.dll",
        "x64\libonigwrap.dll",
        "runtimes\win-x64\native\WebView2Loader.dll",
        "zh-CN\TeXShift.AddIn.resources.dll",
        "zh-CN\TeXShift.Core.resources.dll",
        "Lib\mathjax\es5\tex-mml-chtml.js",
        "Lib\mermaid\mermaid.min.js"
    )

    $missing = @($requiredFiles | Where-Object {
        -not (Test-Path -LiteralPath (Join-Path $StageRoot $_) -PathType Leaf)
    })
    if ($missing.Count -gt 0) {
        throw "Staged payload is incomplete: $($missing -join ', ')"
    }

    foreach ($unexpected in @("arm64", "x86")) {
        if (Test-Path -LiteralPath (Join-Path $StageRoot $unexpected)) {
            throw "Unexpected architecture in staged payload: $unexpected"
        }
    }

    foreach ($unexpected in @(
        "Interop.Microsoft.Office.Interop.OneNote.dll",
        "Microsoft.VisualStudio.Interop.dll"
    )) {
        if (Test-Path -LiteralPath (Join-Path $StageRoot $unexpected) -PathType Leaf) {
            throw "Unexpected stale dependency in staged payload: $unexpected"
        }
    }

    $mathJaxFileCount = (Get-ChildItem -LiteralPath (Join-Path $StageRoot "Lib\mathjax") -Recurse -File).Count
    if ($mathJaxFileCount -lt 100) {
        throw "MathJax payload is incomplete: $mathJaxFileCount files"
    }
}

function Write-PayloadManifest {
    New-Item -ItemType Directory -Path $DiagnosticsRoot -Force | Out-Null
    $manifestPath = Join-Path $DiagnosticsRoot "payload-manifest.json"
    $stagePrefixLength = [IO.Path]::GetFullPath($StageRoot).TrimEnd('\').Length + 1
    $files = Get-ChildItem -LiteralPath $StageRoot -Recurse -File | Sort-Object FullName | ForEach-Object {
        [ordered]@{
            path = $_.FullName.Substring($stagePrefixLength).Replace('\', '/')
            size = $_.Length
            sha256 = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
        }
    }
    $files | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath $manifestPath -Encoding UTF8
    Write-Host "Payload manifest: $manifestPath"
}

function Invoke-Stage {
    $outputRoot = Join-Path $RepoRoot "src\TeXShift.AddIn\bin\x64\$Configuration\net48"
    if (-not (Test-Path -LiteralPath $outputRoot -PathType Container)) {
        throw "AddIn output not found: $outputRoot"
    }

    Reset-ArtifactDirectory $DiagnosticsRoot
    Reset-ArtifactDirectory $StageRoot
    Get-ChildItem -LiteralPath $outputRoot -File | Where-Object {
        $_.Extension -in @(".dll", ".config") -or
        ($Configuration -eq "Debug" -and $_.Extension -eq ".pdb")
    } | Copy-Item -Destination $StageRoot -Force

    foreach ($directory in @("Lib", "x64", "runtimes\win-x64\native", "zh-CN")) {
        Copy-StageDirectory $outputRoot $directory
    }

    Test-StagedPayload
    Write-PayloadManifest
    Write-Host "Staged payload: $StageRoot"
}

function Find-InnoSetupCompiler {
    if ($InnoSetupCompiler) {
        return [IO.Path]::GetFullPath($InnoSetupCompiler)
    }
    if ($env:ISCC_PATH) {
        return [IO.Path]::GetFullPath($env:ISCC_PATH)
    }

    $command = Get-Command ISCC.exe -ErrorAction SilentlyContinue
    if ($command) {
        return $command.Source
    }

    $candidates = @(
        (Join-Path $env:ProgramFiles "Inno Setup 7\ISCC.exe"),
        (Join-Path ${env:ProgramFiles(x86)} "Inno Setup 7\ISCC.exe"),
        (Join-Path $env:LOCALAPPDATA "Programs\Inno Setup 7\ISCC.exe")
    )
    return $candidates | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } | Select-Object -First 1
}

function Get-ReleaseAssemblyMetadata {
    param([string]$Path)

    $inspectionScript = @'
$ErrorActionPreference = "Stop"
$path = [Environment]::GetEnvironmentVariable("TEXSHIFT_INSPECTION_ASSEMBLY")
$assemblyName = [Reflection.AssemblyName]::GetAssemblyName($path)
$publicKeyToken = -join ($assemblyName.GetPublicKeyToken() | ForEach-Object { $_.ToString("x2") })
$assembly = [Reflection.Assembly]::LoadFrom($path)
$connectType = $assembly.GetType("TeXShift.AddIn.Connect", $true)
$attributes = $connectType.GetCustomAttributesData()
$clsid = [string](($attributes |
    Where-Object { $_.AttributeType.FullName -eq "System.Runtime.InteropServices.GuidAttribute" } |
    Select-Object -First 1).ConstructorArguments[0].Value)
$progId = [string](($attributes |
    Where-Object { $_.AttributeType.FullName -eq "System.Runtime.InteropServices.ProgIdAttribute" } |
    Select-Object -First 1).ConstructorArguments[0].Value)
$peKind = [Reflection.PortableExecutableKinds]::NotAPortableExecutableImage
$machine = [Reflection.ImageFileMachine]::I386
$assembly.ManifestModule.GetPEKind([ref]$peKind, [ref]$machine)
[ordered]@{
    assemblyVersion = $assemblyName.Version.ToString()
    publicKeyToken = $publicKeyToken
    clsid = $clsid
    progId = $progId
    machine = $machine.ToString()
} | ConvertTo-Json -Compress
'@
    $encodedScript = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($inspectionScript))
    $previousPath = [Environment]::GetEnvironmentVariable("TEXSHIFT_INSPECTION_ASSEMBLY", "Process")
    try {
        [Environment]::SetEnvironmentVariable("TEXSHIFT_INSPECTION_ASSEMBLY", $Path, "Process")
        $output = @(& powershell.exe -NoProfile -NonInteractive -EncodedCommand $encodedScript 2>&1)
        if ($LASTEXITCODE -ne 0) {
            throw "Release assembly inspection failed: $($output -join [Environment]::NewLine)"
        }
    }
    finally {
        [Environment]::SetEnvironmentVariable("TEXSHIFT_INSPECTION_ASSEMBLY", $previousPath, "Process")
    }

    $json = $output | Where-Object { $_ -is [string] -and $_.TrimStart().StartsWith("{") } | Select-Object -Last 1
    if (-not $json) {
        throw "Release assembly inspection returned no metadata."
    }
    return $json | ConvertFrom-Json
}

function Invoke-Package {
    if ($Configuration -ne "Release") {
        throw "Package requires -Configuration Release."
    }
    if (-not $Version) {
        throw "Package requires -Version <major.minor.patch>."
    }
    if (-not (Test-Path -LiteralPath $InstallerScript -PathType Leaf)) {
        throw "Installer script not found: $InstallerScript"
    }

    $assemblyPath = Join-Path $StageRoot "TeXShift.AddIn.dll"
    $metadata = Get-ReleaseAssemblyMetadata $assemblyPath
    $publicKeyToken = [string]$metadata.publicKeyToken
    if (-not $publicKeyToken) {
        throw "Release payload is not strong-name signed."
    }

    $clsid = [string]$metadata.clsid
    $progId = [string]$metadata.progId
    if ($clsid -ne "1EE8F914-ECBD-4709-92C0-E770851C4966" -or $progId -ne "TeXShift.AddIn.Connect") {
        throw "Release payload has the wrong COM identity: $clsid / $progId"
    }

    if ([string]$metadata.machine -ne "AMD64") {
        throw "Release payload is not x64: $($metadata.machine)"
    }

    $compiler = Find-InnoSetupCompiler
    if (-not $compiler) {
        throw "ISCC.exe not found. Install Inno Setup 7 or pass -InnoSetupCompiler."
    }

    $arguments = @(
        "/Qp",
        "/DAppVersion=$Version",
        "/DSourceDir=$StageRoot",
        "/DOutputDir=$PackageRoot",
        "/DAssemblyVersion=$($metadata.assemblyVersion)",
        "/DPublicKeyToken=$publicKeyToken",
        $InstallerScript
    )
    Write-Host ("ISCC " + ($arguments -join " "))
    try {
        & $compiler @arguments
        if ($LASTEXITCODE -ne 0) {
            throw "Inno Setup failed with exit code $LASTEXITCODE."
        }

        $installer = Join-Path $PackageRoot "TeXShift-$Version-x64-Setup.exe"
        if (-not (Test-Path -LiteralPath $installer -PathType Leaf)) {
            throw "Installer output not found: $installer"
        }

        $hash = (Get-FileHash -LiteralPath $installer -Algorithm SHA256).Hash.ToLowerInvariant()
        $hashPath = "$installer.sha256"
        Set-Content -LiteralPath $hashPath -Value "$hash  $([IO.Path]::GetFileName($installer))" -Encoding ASCII

        $signature = Get-AuthenticodeSignature -LiteralPath $installer
        Write-Host "Installer: $installer"
        Write-Host "SHA-256: $hash"
        Write-Host "Authenticode: $($signature.Status)"
    }
    catch {
        Reset-ArtifactDirectory $PackageRoot
        throw
    }
}

function Invoke-E2E {
    if (-not $E2EInput -or -not $E2EOutput) {
        throw "E2E requires -E2EInput and -E2EOutput."
    }

    $runner = Join-Path $RepoRoot "tests\TeXShift.Tests.E2E\bin\x64\$Configuration\net48\TeXShift.Tests.E2E.exe"
    $inputFlag = if ($E2ECommand -eq "reverse-xml") { "--input-xml" } else { "--input" }
    & $runner $E2ECommand $inputFlag $E2EInput --output $E2EOutput
    if ($LASTEXITCODE -ne 0) {
        throw "E2E failed with exit code $LASTEXITCODE."
    }
}

switch ($Target) {
    "Clean" {
        Invoke-Clean
    }
    "Restore" {
        Invoke-Restore
    }
    "Build" {
        Invoke-Restore
        Invoke-Build
    }
    "Test" {
        Invoke-Restore
        Invoke-Build
        Invoke-UnitTests
    }
    "Stage" {
        Invoke-Clean
        Invoke-Restore
        Invoke-Build
        Invoke-Stage
    }
    "Verify" {
        Invoke-Clean
        Invoke-Restore
        Invoke-Build
        Invoke-UnitTests
        Invoke-Stage
    }
    "Package" {
        Reset-ArtifactDirectory $PackageRoot
        Invoke-Clean
        Invoke-Restore
        Invoke-Build
        Invoke-UnitTests
        Invoke-Stage
        Invoke-Package
    }
    "E2E" {
        Invoke-Clean
        Invoke-Restore
        Invoke-Build
        Invoke-E2E
    }
    "CI" {
        Invoke-Clean
        Invoke-Restore -Locked
        Invoke-Build
        Invoke-UnitTests
        Invoke-Stage
    }
    "All" {
        if ($Version) {
            Reset-ArtifactDirectory $PackageRoot
        }
        Invoke-Clean
        Invoke-Restore
        Invoke-Build
        Invoke-UnitTests
        Invoke-Stage
        if ($Version) {
            Invoke-Package
        }
    }
}
