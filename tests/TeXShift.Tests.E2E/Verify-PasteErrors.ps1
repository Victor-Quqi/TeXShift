[CmdletBinding()]
param(
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Debug',
    [string]$OutputDirectory
)

$ErrorActionPreference = 'Stop'
$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '../..')).Path
$runner = Join-Path $PSScriptRoot "bin/x64/$Configuration/net48/TeXShift.Tests.E2E.exe"
if (-not $OutputDirectory) {
    $OutputDirectory = Join-Path $repoRoot ('TestResults/paste-errors-' + [Guid]::NewGuid().ToString('N'))
}
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$inputFile = Join-Path $OutputDirectory 'clipboard.txt'
try {
    foreach ($case in @(@{ Name = 'empty'; Text = '' }, @{ Name = 'whitespace'; Text = " `r`n`t" })) {
        [IO.File]::WriteAllText($inputFile, $case.Text)
        $output = Join-Path $OutputDirectory $case.Name
        New-Item -ItemType Directory -Path $output -Force | Out-Null
        $json = & $runner paste-convert --markdown 'Retained body' --paste-input $inputFile --output $output 2> (Join-Path $output 'stderr.txt')
        $code = $LASTEXITCODE
        $result = ($json -join "`n") | ConvertFrom-Json
        if ($code -eq 0 -or $result.cleanupStatus -ne 'passed') {
            throw "$($case.Name): expected a rejected paste with successful cleanup. $($result.error)"
        }
        [xml]$dialog = [IO.File]::ReadAllText((Join-Path $output 'F11_Error_Dialog.txt'))
        if ($dialog.UI.id -ne 'TeXShift.ErrorDialog' -or
            -not $dialog.SelectSingleNode('//Control[@id="DetailsToggle"]') -or
            -not $dialog.SelectSingleNode('//Control[@type="ControlType.Button" and (@name="复制详情" or @name="Copy Details")]') -or
            -not $dialog.SelectSingleNode('//Control[@type="ControlType.Button" and (@name="打开调试文件夹" or @name="Open Debug Folder")]') -or
            -not $dialog.SelectSingleNode('//Control[@type="ControlType.Button" and (@name="确定" or @name="OK")]')) {
            throw "$($case.Name): expected the original error dialog with details and action buttons."
        }
        if (-not $dialog.SelectSingleNode('//Control[@type="ControlType.Text" and (@name="请先复制 Markdown 文本。" or @name="Copy Markdown text first.")]')) {
            throw "$($case.Name): missing clipboard validation message."
        }
        if ($dialog.UI.ownedByOneNote -ne 'true') { throw "$($case.Name): error dialog lost its OneNote owner." }
        if ([long]$dialog.UI.observedAfterMs -ge 1000) {
            throw "$($case.Name): error dialog took $($dialog.UI.observedAfterMs) ms to appear."
        }
        [xml]$original = [IO.File]::ReadAllText((Join-Path $output $result.files.originalXml))
        [xml]$after = [IO.File]::ReadAllText((Join-Path $output 'F08_Failed_Selection.xml'))
        foreach ($document in @($original, $after)) {
            foreach ($attribute in @($document.SelectNodes('//@selected'))) {
                [void]$attribute.OwnerElement.RemoveAttributeNode($attribute)
            }
        }
        $beforeOutlines = @($original.SelectNodes('//*[local-name()="Outline"]') | ForEach-Object OuterXml)
        $afterOutlines = @($after.SelectNodes('//*[local-name()="Outline"]') | ForEach-Object OuterXml)
        if (Compare-Object $beforeOutlines $afterOutlines) { throw "$($case.Name): rejected paste changed the page." }
        Write-Host "$($case.Name): PASS (dialog $($dialog.UI.observedAfterMs) ms)"
    }
}
finally {
    if (Test-Path -LiteralPath $inputFile) { Remove-Item -LiteralPath $inputFile }
}
