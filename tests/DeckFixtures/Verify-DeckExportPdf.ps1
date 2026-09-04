$ErrorActionPreference = 'Stop'

$project = Join-Path $PSScriptRoot '..\..\src\officecli\officecli.csproj'
$fixture = Join-Path $PSScriptRoot 'technology-dark.workmate-deck.json'
if (-not (Test-Path -LiteralPath $fixture)) { throw "Missing fixture: $fixture" }

$tmpDir = Join-Path ([System.IO.Path]::GetTempPath()) 'officecli-deck-export-pdf'
[System.IO.Directory]::CreateDirectory($tmpDir) | Out-Null

# --- missing exporter → exporter_not_found (documents plugin / LibreOffice dependency) ---
$env:OFFICECLI_PLUGIN_EXPORTER_PDF = $null
Remove-Item Env:OFFICECLI_PLUGIN_EXPORTER_PDF -ErrorAction SilentlyContinue
$missingOut = Join-Path $tmpDir 'missing.pdf'
$missingJson = & dotnet run --project $project -c Release --no-build -- deck export-pdf $fixture -o $missingOut --json 2>&1 | Out-String
if ($LASTEXITCODE -eq 0) { throw "Expected exporter_not_found when no PDF plugin is installed`n$missingJson" }
if ($missingJson -notmatch 'exporter_not_found') { throw "Expected exporter_not_found in error envelope`n$missingJson" }
Write-Output 'export-pdf missing-exporter path ok (exporter_not_found)'

# --- stub exporter via OFFICECLI_PLUGIN_EXPORTER_PDF ---
$stubProj = Join-Path $PSScriptRoot 'FakePdfExporter\FakePdfExporter.csproj'
$stubOut = Join-Path $tmpDir 'FakePdfExporter'
& dotnet build $stubProj -c Release -o $stubOut
if ($LASTEXITCODE -ne 0) { throw 'Failed to build FakePdfExporter stub' }
$stubExe = Join-Path $stubOut $(if ($IsWindows) { 'FakePdfExporter.exe' } else { 'FakePdfExporter' })
if (-not (Test-Path -LiteralPath $stubExe)) { throw "Stub exe missing: $stubExe" }
$env:OFFICECLI_PLUGIN_EXPORTER_PDF = $stubExe

$pdf = Join-Path $tmpDir 'technology-dark.pdf'
$pptx = Join-Path $tmpDir 'technology-dark.pptx'
$json = & dotnet run --project $project -c Release --no-build -- deck export-pdf $fixture -o $pdf --pptx $pptx --json
if ($LASTEXITCODE -ne 0) { throw "export-pdf failed`n$json" }
$obj = $json | ConvertFrom-Json
if ($obj.success -ne $true) { throw "Expected success=true; got $json" }
if (-not (Test-Path -LiteralPath $obj.output)) { throw "PDF missing: $($obj.output)" }
if (-not (Test-Path -LiteralPath $pptx)) { throw "Kept PPTX missing: $pptx" }
if ($obj.plugin -ne 'officecli-pdf-stub') { throw "Expected plugin officecli-pdf-stub; got $($obj.plugin)" }
$pdfBytes = [System.IO.File]::ReadAllBytes($obj.output)
$header = [System.Text.Encoding]::ASCII.GetString($pdfBytes[0..([Math]::Min(7, $pdfBytes.Length - 1))])
if (-not $header.StartsWith('%PDF-')) { throw "PDF header missing: $header" }

# Default output path: sibling of a copied spec (keep relative assets)
$specCopy = Join-Path $tmpDir 'sample.workmate-deck.json'
Copy-Item -LiteralPath $fixture -Destination $specCopy -Force
$asset = Join-Path $PSScriptRoot 'visual.svg'
if (Test-Path -LiteralPath $asset) {
    Copy-Item -LiteralPath $asset -Destination (Join-Path $tmpDir 'visual.svg') -Force
}
$defaultPdf = Join-Path $tmpDir 'sample.pdf'
if (Test-Path -LiteralPath $defaultPdf) { Remove-Item -LiteralPath $defaultPdf -Force }
$json2 = & dotnet run --project $project -c Release --no-build -- deck export-pdf $specCopy --json
if ($LASTEXITCODE -ne 0) { throw "export-pdf default -o failed`n$json2" }
$obj2 = $json2 | ConvertFrom-Json
if (-not (Test-Path -LiteralPath $defaultPdf)) { throw "Default PDF path not written: $defaultPdf" }
if ([System.IO.Path]::GetFullPath($obj2.output) -ne [System.IO.Path]::GetFullPath($defaultPdf)) {
    throw "Default output mismatch: $($obj2.output) vs $defaultPdf"
}

Remove-Item Env:OFFICECLI_PLUGIN_EXPORTER_PDF -ErrorAction SilentlyContinue
Write-Output 'Verify-DeckExportPdf passed'
