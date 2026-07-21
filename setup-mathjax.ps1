[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

# A PSModulePath inherited from a PowerShell 7 host shadows Windows PowerShell's
# built-in modules and breaks cmdlet autoloading (e.g. Get-FileHash); make sure
# $PSHOME\Modules is searched first.
if ($PSVersionTable.PSEdition -ne "Core") {
    $env:PSModulePath = (Join-Path $PSHome "Modules") + ";" + $env:PSModulePath
}

$Version = "3.2.2"
$ArchiveUrl = "https://registry.npmjs.org/mathjax/-/mathjax-$Version.tgz"
$ExpectedSha512 = "Bt+SSVU8eBG27zChVewOicYs7Xsdt40qm4+UpHyX7k0/O9NliPc+x77k1/FEsPsjKPZGJvtRZM1vO+geW0OhGw=="
$RepoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$MathJaxPath = [IO.Path]::GetFullPath((Join-Path $RepoRoot "src\TeXShift.AddIn\Lib\mathjax"))
$VersionFile = Join-Path $MathJaxPath ".texshift-version"
$RequiredFile = Join-Path $MathJaxPath "es5\tex-mml-chtml.js"
$ArtifactsRoot = [IO.Path]::GetFullPath((Join-Path $RepoRoot "artifacts"))
$TempRoot = Join-Path $ArtifactsRoot ("temp\mathjax-" + [Guid]::NewGuid().ToString("N"))

if ((Test-Path -LiteralPath $RequiredFile -PathType Leaf) -and
    (Test-Path -LiteralPath $VersionFile -PathType Leaf) -and
    ((Get-Content -LiteralPath $VersionFile -Raw).Trim() -eq $Version)) {
    Write-Host "MathJax $Version is ready."
    exit 0
}

$repoPrefix = [IO.Path]::GetFullPath($RepoRoot).TrimEnd('\') + '\'
if (-not $MathJaxPath.StartsWith($repoPrefix, [StringComparison]::OrdinalIgnoreCase)) {
    throw "Refusing to modify a MathJax path outside the repository: $MathJaxPath"
}

New-Item -ItemType Directory -Path $TempRoot -Force | Out-Null
$archivePath = Join-Path $TempRoot "mathjax.tgz"
$extractPath = Join-Path $TempRoot "extract"
New-Item -ItemType Directory -Path $extractPath | Out-Null

try {
    Write-Host "Downloading MathJax $Version..."
    Invoke-WebRequest -Uri $ArchiveUrl -OutFile $archivePath -UseBasicParsing

    $sha512 = [Security.Cryptography.SHA512]::Create()
    try {
        $actualSha512 = [Convert]::ToBase64String(
            $sha512.ComputeHash([IO.File]::ReadAllBytes($archivePath)))
    }
    finally {
        $sha512.Dispose()
    }

    if ($actualSha512 -ne $ExpectedSha512) {
        throw "MathJax archive checksum mismatch."
    }

    & tar.exe -xf $archivePath -C $extractPath
    if ($LASTEXITCODE -ne 0) {
        throw "Unable to extract the MathJax archive."
    }

    $sourcePath = Join-Path $extractPath "package\es5"
    if (-not (Test-Path -LiteralPath (Join-Path $sourcePath "tex-mml-chtml.js") -PathType Leaf)) {
        throw "The MathJax archive does not contain the expected es5 payload."
    }

    if (Test-Path -LiteralPath $MathJaxPath) {
        Remove-Item -LiteralPath $MathJaxPath -Recurse -Force
    }
    New-Item -ItemType Directory -Path $MathJaxPath | Out-Null
    Copy-Item -LiteralPath $sourcePath -Destination $MathJaxPath -Recurse -Force
    Set-Content -LiteralPath $VersionFile -Value $Version -Encoding ASCII
    Write-Host "MathJax $Version restored."
}
finally {
    if (Test-Path -LiteralPath $TempRoot) {
        Remove-Item -LiteralPath $TempRoot -Recurse -Force
    }
}
