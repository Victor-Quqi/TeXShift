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
    if (-not $targetAfter) { throw "$Name lost the target outline." }
    $selectedTarget = $selection.SelectSingleNode('//*[local-name()="Outline" and @selected]')
    if (-not $selectedTarget -or $selectedTarget.objectID -ne $target.objectID -or
        $selection.SelectSingleNode('//*[local-name()="Title"]//*[@selected]')) {
        throw "$Name left the caret outside the insertion target."
    }
    foreach ($axis in @('x', 'y')) {
        if ($target.SelectSingleNode('*[local-name()="Position"]').GetAttribute($axis) -ne
            $targetAfter.SelectSingleNode('*[local-name()="Position"]').GetAttribute($axis)) {
            throw "$Name moved the insertion position."
        }
    }
    Write-Host "${Name}: PASS"
    return @{ Original = $original; Final = $final; Target = $targetAfter }
}

function Assert-RetainedLeadingParagraph {
    param($Result, [string]$Block)
    $original = $Result.Original.SelectSingleNode('//*[local-name()="Outline" and @selected]/*[local-name()="OEChildren"]/*[1]')
    $paragraphs = $Result.Target.SelectNodes('*[local-name()="OEChildren"]/*')
    if ($paragraphs.Count -lt 3 -or $paragraphs[0].objectID -ne $original.objectID -or
        -not $paragraphs[0].SelectSingleNode('*[local-name()="T"]') -or $paragraphs[0].InnerText -ne '' -or
        -not $paragraphs[1].SelectSingleNode("*[local-name()='$Block']")) {
        throw "$Block paste must retain the original empty paragraph immediately before the inserted block."
    }
}

$inline = Join-Path $repoRoot 'tests/TestData/paste-inline.md'
$blocks = Join-Path $repoRoot 'tests/TestData/paste-convert.md'
$imagePath = Join-Path $repoRoot 'tests/TestData/Assets/TeXShift_512.png'
$imagePage = Join-Path $OutputDirectory 'image-page.xml'
$tableInput = Join-Path $OutputDirectory 'table.md'
$imageInput = Join-Path $OutputDirectory 'image.md'
$nestedPage = Join-Path $OutputDirectory 'nested-page.xml'
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
    Assert-RetainedLeadingParagraph $table 'Table'
    $image = Invoke-PasteCase 'empty-image-first' @('--markdown', 'Original body', '--empty-outline', '--paste-input', $imageInput)
    Assert-RetainedLeadingParagraph $image 'Image'

    $selected = Invoke-PasteCase 'selected-table' @('--markdown', 'Replace me', '--select', 'Replace me', '--paste-input', $tableInput)
    Assert-RetainedLeadingParagraph $selected 'Table'
    if ($selected.Final.DocumentElement.InnerText.Contains('Replace me')) { throw 'Paste did not replace the selected text.' }
    $existing = Invoke-PasteCase 'existing-outline-table' @('--markdown', 'Original body', '--paste-input', $tableInput)
    Assert-RetainedLeadingParagraph $existing 'Table'
    if (-not $existing.Final.DocumentElement.InnerText.Contains('Original body')) { throw 'Paste lost surrounding text.' }

    [IO.File]::WriteAllText($nestedPage, @"
<one:Page xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote">
  <one:Outline><one:Position x="36" y="90" z="0"/><one:OEChildren>
    <one:OE><one:List><one:Bullet bullet="2"/></one:List><one:T><![CDATA[Parent bullet]]></one:T>
      <one:OEChildren><one:OE><one:List><one:Bullet bullet="2"/></one:List><one:T><![CDATA[Child bullet]]></one:T></one:OE></one:OEChildren>
    </one:OE>
  </one:OEChildren></one:Outline>
</one:Page>
"@)
    $undoArgs = if ($Driver -eq 'ribbon') { @('--verify-undo') } else { @() }
    $nested = Invoke-PasteCase 'nested-list-table' (@('--page', $nestedPage, '--caret', 'Parent bullet', '--paste-input', $tableInput) + $undoArgs)
    $childBefore = $nested.Original.SelectSingleNode('//*[local-name()="T" and text()="Child bullet"]/..')
    $childAfter = $nested.Final.SelectSingleNode("//*[@objectID='$($childBefore.objectID)']")
    if (-not $childAfter -or $childBefore.ParentNode.ParentNode.objectID -ne $childAfter.ParentNode.ParentNode.objectID) {
        throw 'Paste moved an existing child to a different parent identity.'
    }
    if (-not $nested.Target.InnerText.Contains('Parent bullet') -or -not $nested.Target.InnerText.Contains('Child bullet') -or
        -not $nested.Target.SelectSingleNode('.//*[local-name()="Table"]')) { throw 'Paste lost nested list content.' }

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
    foreach ($path in @($imagePage, $tableInput, $imageInput, $nestedPage)) {
        if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path }
    }
}
