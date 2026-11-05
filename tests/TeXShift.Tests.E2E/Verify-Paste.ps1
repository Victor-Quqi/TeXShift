[CmdletBinding()]
param(
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Debug',
    [ValidateSet('ribbon', 'core')]
    [string]$Driver = 'ribbon',
    [string]$OutputDirectory
)

$ErrorActionPreference = 'Stop'
$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '../..')).Path
$runner = Join-Path $PSScriptRoot "bin/x64/$Configuration/net48/TeXShift.Tests.E2E.exe"
if (-not $OutputDirectory) {
    $OutputDirectory = Join-Path $repoRoot ('TestResults/paste-' + [Guid]::NewGuid().ToString('N'))
}
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null

function Invoke-PasteCase {
    param([string]$Name, [string[]]$Arguments)
    $output = Join-Path $OutputDirectory $Name
    New-Item -ItemType Directory -Path $output -Force | Out-Null
    $json = & $runner paste-convert @Arguments --driver $Driver --output $output 2> (Join-Path $output 'stderr.txt')
    $code = $LASTEXITCODE
    $result = ($json -join "`n") | ConvertFrom-Json
    if ($code -ne 0 -or $result.cleanupStatus -ne 'passed') {
        throw "$Name failed: exit=$code cleanup=$($result.cleanupStatus) $($result.error)"
    }
    [xml]$original = [IO.File]::ReadAllText((Join-Path $output $result.files.originalXml))
    [xml]$final = [IO.File]::ReadAllText((Join-Path $output $result.files.finalXml))
    [xml]$selection = [IO.File]::ReadAllText((Join-Path $output $result.files.finalSelectionXml))
    $ids = @($final.SelectNodes('//@objectID') | ForEach-Object Value)
    if (@($ids | Sort-Object -Unique).Count -ne $ids.Count) { throw "$Name duplicated an object ID." }
    $target = $original.SelectSingleNode('//*[local-name()="Outline" and @selected]')
    if (-not $target) { throw "$Name had no insertion target." }
    $targetAfter = $final.SelectSingleNode("//*[@objectID='$($target.objectID)']")
    if (-not $targetAfter -and $Name.StartsWith('empty-')) {
        $targetAfter = $selection.SelectSingleNode('//*[local-name()="Outline" and @selected]')
    }
    if (-not $targetAfter) { throw "$Name lost the target outline." }
    foreach ($axis in @('x', 'y')) {
        if ($target.SelectSingleNode('*[local-name()="Position"]').GetAttribute($axis) -ne
            $targetAfter.SelectSingleNode('*[local-name()="Position"]').GetAttribute($axis)) {
            throw "$Name moved the insertion position."
        }
    }
    Write-Host "${Name}: PASS"
    return @{ Original = $original; Final = $final; Target = $targetAfter }
}

$inline = Join-Path $repoRoot 'tests/TestData/paste-inline.md'
$blocks = Join-Path $repoRoot 'tests/TestData/paste-convert.md'
$imagePath = Join-Path $repoRoot 'tests/TestData/Assets/TeXShift_512.png'
$imagePage = Join-Path $OutputDirectory 'image-page.xml'
$tableInput = Join-Path $OutputDirectory 'table.md'
$imageInput = Join-Path $OutputDirectory 'image.md'
try {
    foreach ($case in @(@{ Name = 'empty-inline'; Input = $inline }, @{ Name = 'empty-block'; Input = $blocks })) {
        $result = Invoke-PasteCase $case.Name @('--markdown', 'Original body', '--empty-outline', '--paste-input', $case.Input)
        if (-not $result.Final.DocumentElement.InnerText.Contains('Original body')) { throw 'Paste changed the other text box.' }
        if ($case.Name -eq 'empty-inline' -and -not $result.Final.DocumentElement.InnerText.Contains('Pasted inline')) {
            throw 'Inline paste produced no text.'
        }
        if ($case.Name -eq 'empty-block' -and (-not $result.Final.SelectSingleNode('//*[local-name()="Table"]') -or
            -not $result.Final.OuterXml.Contains('[if mathML]'))) { throw 'Block paste lost its table or formula.' }
    }
    [IO.File]::WriteAllText($tableInput, "| A | B |`n|---|---|`n| 1 | 2 |")
    [IO.File]::WriteAllText($imageInput, "![sample]($($imagePath.Replace('\', '/')))")
    $table = Invoke-PasteCase 'empty-table-first' @('--markdown', 'Original body', '--empty-outline', '--paste-input', $tableInput)
    if (-not $table.Target.SelectSingleNode('*[local-name()="OEChildren"]/*[1]/*[local-name()="Table"]')) {
        throw 'Table paste must start with the table, without a leading empty paragraph.'
    }
    $image = Invoke-PasteCase 'empty-image-first' @('--markdown', 'Original body', '--empty-outline', '--paste-input', $imageInput)
    if (-not $image.Target.SelectSingleNode('*[local-name()="OEChildren"]/*[1]/*[local-name()="Image"]')) {
        throw 'Image paste must start with the image, without a leading empty paragraph.'
    }

    $selected = Invoke-PasteCase 'selected-table' @('--markdown', 'Replace me', '--select', 'Replace me', '--paste-input', $tableInput)
    if ($selected.Final.DocumentElement.InnerText.Contains('Replace me') -or
        -not $selected.Target.SelectSingleNode('*[local-name()="OEChildren"]/*[1]/*[local-name()="Table"]')) {
        throw 'Replacing text with a table left an extra paragraph.'
    }
    $existing = Invoke-PasteCase 'existing-outline-table' @('--markdown', 'Original body', '--paste-input', $tableInput)
    if (-not $existing.Final.DocumentElement.InnerText.Contains('Original body') -or
        -not $existing.Target.SelectSingleNode('*[local-name()="OEChildren"]/*[1]/*[local-name()="Table"]')) {
        throw 'Pasting into an existing outline added a leading paragraph or lost surrounding text.'
    }

    $escapedImage = [System.Security.SecurityElement]::Escape($imagePath)
    [IO.File]::WriteAllText($imagePage, @"
<one:Page xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote">
  <one:Outline><one:Position x="36" y="90" z="0"/><one:OEChildren>
    <one:OE><one:T><![CDATA[Before]]></one:T></one:OE>
    <one:OE><one:Image format="png"><one:Size width="64" height="64"/><one:File path="$escapedImage"/></one:Image></one:OE>
  </one:OEChildren></one:Outline>
  <one:Outline><one:Position x="300" y="90" z="1"/><one:OEChildren>
    <one:OE><one:T><![CDATA[Other outline]]></one:T></one:OE>
    <one:OE><one:Image format="png"><one:Size width="64" height="64"/><one:File path="$escapedImage"/></one:Image></one:OE>
  </one:OEChildren></one:Outline>
</one:Page>
"@)
    $result = Invoke-PasteCase 'retained-images' @('--page', $imagePage, '--caret', 'Before', '--paste-input', $inline)
    $before = @($result.Original.SelectNodes('//*[local-name()="Image"]/*[local-name()="CallbackID"]/@callbackID') | ForEach-Object Value)
    $after = @($result.Final.SelectNodes('//*[local-name()="Image"]/*[local-name()="CallbackID"]/@callbackID') | ForEach-Object Value)
    if ($before.Count -ne 2 -or $after.Count -ne 2 -or (Compare-Object $before $after)) {
        throw 'Paste lost an existing image.'
    }
    if (-not $result.Final.DocumentElement.InnerText.Contains('Other outline')) { throw 'Paste changed the other outline.' }
}
finally {
    foreach ($path in @($imagePage, $tableInput, $imageInput)) {
        if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path }
    }
}
