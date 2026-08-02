param(
  [string]$SourceDir = (Join-Path $PSScriptRoot 'OfficialDocumentFormat_0.01'),
  [string]$AddinName = 'OfficialDocumentFormat',
  [string]$Version = '0.01'
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $SourceDir)) {
  throw "Add-in source directory not found: $SourceDir"
}

$addinRoot = Join-Path $env:APPDATA 'kingsoft\wps\jsaddons'
$targetDir = Join-Path $addinRoot ("{0}_{1}" -f $AddinName, $Version)
$publishPath = Join-Path $addinRoot 'publish.xml'

New-Item -ItemType Directory -Force -Path $addinRoot | Out-Null
New-Item -ItemType Directory -Force -Path $targetDir | Out-Null
Copy-Item -LiteralPath (Join-Path $SourceDir 'main.js') -Destination (Join-Path $targetDir 'main.js') -Force
Copy-Item -LiteralPath (Join-Path $SourceDir 'ribbon.xml') -Destination (Join-Path $targetDir 'ribbon.xml') -Force

$fileUrl = ('file:///' + ($targetDir -replace '\\', '/'))
$publishXml = @"
<?xml version="1.0" encoding="UTF-8"?>
<jsplugins>
  <jsplugin name="$AddinName" type="w" version="$Version" url="$fileUrl" enable="enable"/>
</jsplugins>
"@

Set-Content -LiteralPath $publishPath -Value $publishXml -Encoding UTF8

Write-Host "Installed WPS add-in to: $targetDir"
Write-Host "Wrote publish.xml: $publishPath"
Write-Host "Restart WPS Writer, then look for the 公文格式 tab."
