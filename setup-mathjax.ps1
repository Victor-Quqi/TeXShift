# Setup MathJax for TeXShift
#
# This script downloads MathJax from npm if it doesn't exist locally.
# Run this script before building the project for the first time.

param(
    [string]$Version = "3"
)

$ErrorActionPreference = "Stop"

$RepoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$MathJaxPath = Join-Path $RepoRoot "src\TeXShift.AddIn\Lib\mathjax"

Write-Host "TeXShift MathJax Setup" -ForegroundColor Cyan
Write-Host "======================" -ForegroundColor Cyan
Write-Host ""

# Check if MathJax already exists
if (Test-Path $MathJaxPath) {
    Write-Host "✓ MathJax already exists at: $MathJaxPath" -ForegroundColor Green
    Write-Host "  Nothing to do." -ForegroundColor Gray
    exit 0
}

Write-Host "MathJax not found. Downloading from npm..." -ForegroundColor Yellow
Write-Host ""

# Check if npm is available
try {
    $npmVersion = & npm --version 2>&1
    Write-Host "✓ npm version: $npmVersion" -ForegroundColor Green
} catch {
    Write-Host "✗ npm is not installed or not in PATH" -ForegroundColor Red
    Write-Host ""
    Write-Host "Please install Node.js from https://nodejs.org/" -ForegroundColor Yellow
    exit 1
}

# Create temp directory
$TempDir = Join-Path $RepoRoot "temp_mathjax_download"
if (Test-Path $TempDir) {
    Remove-Item -Recurse -Force $TempDir
}
New-Item -ItemType Directory -Path $TempDir | Out-Null

try {
    Write-Host "Downloading MathJax@$Version..." -ForegroundColor Cyan

    # Download MathJax using npm
    Push-Location $TempDir
    & npm install "mathjax@$Version" --silent --no-audit --no-fund 2>&1 | Out-Null
    Pop-Location

    if ($LASTEXITCODE -ne 0) {
        throw "npm install failed with exit code $LASTEXITCODE"
    }

    # Copy only the es5 folder we need
    $SourcePath = Join-Path $TempDir "node_modules\mathjax\es5"
    if (-not (Test-Path $SourcePath)) {
        throw "MathJax es5 folder not found at: $SourcePath"
    }

    Write-Host "Copying MathJax files..." -ForegroundColor Cyan

    # Create target directory
    $LibPath = Join-Path $RepoRoot "src\TeXShift.AddIn\Lib"
    if (-not (Test-Path $LibPath)) {
        New-Item -ItemType Directory -Path $LibPath | Out-Null
    }

    Copy-Item -Recurse $SourcePath $MathJaxPath

    Write-Host ""
    Write-Host "✓ MathJax successfully installed!" -ForegroundColor Green
    Write-Host "  Location: $MathJaxPath" -ForegroundColor Gray

    # Get file count
    $fileCount = (Get-ChildItem -Recurse $MathJaxPath -File | Measure-Object).Count
    Write-Host "  Files: $fileCount" -ForegroundColor Gray

} catch {
    Write-Host ""
    Write-Host "✗ Error: $_" -ForegroundColor Red
    exit 1
} finally {
    # Cleanup temp directory
    if (Test-Path $TempDir) {
        Remove-Item -Recurse -Force $TempDir
    }
}

Write-Host ""
Write-Host "You can now build the project." -ForegroundColor Cyan
