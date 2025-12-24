param(
    [ValidateSet("Build", "Test", "All")]
    [string]$Target = "Build",
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Debug"
)

$ErrorActionPreference = "Stop"

$RepoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$CoreProject = Join-Path $RepoRoot "src\TeXShift.Core\TeXShift.Core.csproj"
$TestProject = Join-Path $RepoRoot "tests\TeXShift.Tests.E2E\TeXShift.Tests.E2E.csproj"

function Invoke-Dotnet {
    param(
        [string[]]$Arguments
    )

    Write-Host ("dotnet " + ($Arguments -join " "))
    & dotnet @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "Command failed with exit code $LASTEXITCODE."
    }
}

function Invoke-Build {
    Invoke-Dotnet @("build", $CoreProject, "-c", $Configuration, "-p:Platform=x64")
    Invoke-Dotnet @("build", $TestProject, "-c", $Configuration, "-p:Platform=x64")
}

function Invoke-Tests {
    param(
        [switch]$NoBuild
    )

    $testArgs = @("test", $TestProject, "-c", $Configuration, "-p:Platform=x64")
    if ($NoBuild) {
        $testArgs += "--no-build"
    }

    Invoke-Dotnet $testArgs
}

switch ($Target) {
    "Build" { Invoke-Build }
    "Test" { Invoke-Tests }
    "All" { Invoke-Build; Invoke-Tests -NoBuild }
}
