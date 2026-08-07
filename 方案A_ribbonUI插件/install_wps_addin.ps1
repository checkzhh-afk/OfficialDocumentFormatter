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
$publishXml = New-Object System.Xml.XmlDocument
$publishXml.PreserveWhitespace = $true
if (Test-Path -LiteralPath $publishPath) {
  try {
    $publishXml.Load($publishPath)
  } catch {
    $publishXml.RemoveAll()
  }
}
if (-not $publishXml.DocumentElement -or $publishXml.DocumentElement.Name -ne 'jsplugins') {
  $publishXml.RemoveAll()
  $declaration = $publishXml.CreateXmlDeclaration('1.0', 'UTF-8', $null)
  [void]$publishXml.AppendChild($declaration)
  [void]$publishXml.AppendChild($publishXml.CreateElement('jsplugins'))
}

$existingNodes = @($publishXml.DocumentElement.SelectNodes('jsplugin'))
foreach ($node in $existingNodes) {
  if ($node.GetAttribute('name') -eq $AddinName) {
    [void]$publishXml.DocumentElement.RemoveChild($node)
  }
}

$pluginNode = $publishXml.CreateElement('jsplugin')
$pluginNode.SetAttribute('name', $AddinName)
$pluginNode.SetAttribute('type', 'w')
$pluginNode.SetAttribute('version', $Version)
$pluginNode.SetAttribute('url', $fileUrl)
$pluginNode.SetAttribute('enable', 'enable')
[void]$publishXml.DocumentElement.AppendChild($pluginNode)

$settings = New-Object System.Xml.XmlWriterSettings
$settings.Encoding = New-Object System.Text.UTF8Encoding($false)
$settings.Indent = $true
$settings.NewLineChars = "`r`n"
$writer = [System.Xml.XmlWriter]::Create($publishPath, $settings)
try {
  $publishXml.Save($writer)
} finally {
  $writer.Close()
}

Write-Host "Installed WPS add-in to: $targetDir"
Write-Host "Updated publish.xml without removing other add-ins: $publishPath"
Write-Host "Restart WPS Writer, then look for the 公文格式 tab."
