$ErrorActionPreference = 'Stop'

$project = Join-Path $PSScriptRoot '..\..\src\officecli\officecli.csproj'
$fixture = Join-Path $PSScriptRoot 'business-light.workmate-deck.json'

$catalogJson = & dotnet run --project $project -c Release --no-build -- deck catalog --json
if ($LASTEXITCODE -ne 0) { throw "catalog failed`n$catalogJson" }
$catalog = $catalogJson | ConvertFrom-Json
$themeIds = @($catalog.themes | ForEach-Object id)
foreach ($id in @('csbu-workmate', 'csbu-workmate-night')) {
    if ($themeIds -notcontains $id) { throw "Expected brand theme $id in catalog; got $($themeIds -join ', ')" }
}
foreach ($id in @('industry-finance', 'industry-consulting', 'industry-tech', 'industry-education')) {
    if ($themeIds -notcontains $id) { throw "Expected industry theme $id in catalog; got $($themeIds -join ', ')" }
}
if ($catalog.version -ne '1.5.3') { throw "Expected catalog 1.5.3; got $($catalog.version)" }
if ($themeIds.Count -lt 18) { throw "Expected >=18 themes; got $($themeIds.Count)" }
$layoutIds = @($catalog.layouts | ForEach-Object id)
if ($layoutIds -notcontains 'heatmap-table') { throw 'Expected heatmap-table layout' }
Write-Output ("catalog ok: themes={0} layouts={1} version={2}" -f $catalog.themes.Count, $catalog.layouts.Count, $catalog.version)

$dry = & dotnet run --project $project -c Release --no-build -- deck theme-remap $fixture --to csbu-workmate-night --json
if ($LASTEXITCODE -ne 0) { throw "theme-remap dry-run failed`n$dry" }
$dryObj = $dry | ConvertFrom-Json
if ($dryObj.applied -ne $false) { throw 'Dry-run should not apply' }
if ($dryObj.report.toThemeId -ne 'csbu-workmate-night') { throw "Unexpected toThemeId $($dryObj.report.toThemeId)" }
if ($dryObj.report.modeChange -ne 'light->dark') { throw "Expected light->dark; got $($dryObj.report.modeChange)" }
if (-not $dryObj.report.slides -or $dryObj.report.slides.Count -lt 1) { throw 'Expected per-slide remap rows' }
Write-Output ("dry-run ok: needsRemapCount={0} slides={1}" -f $dryObj.report.needsRemapCount, $dryObj.report.slides.Count)

$tmpDir = Join-Path ([System.IO.Path]::GetTempPath()) 'officecli-theme-remap-fixture'
[System.IO.Directory]::CreateDirectory($tmpDir) | Out-Null
$tmp = Join-Path $tmpDir 'business-light.workmate-deck.json'
Copy-Item -LiteralPath $fixture -Destination $tmp -Force
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'visual.svg') -Destination (Join-Path $tmpDir 'visual.svg') -Force
$apply = & dotnet run --project $project -c Release --no-build -- deck theme-remap $tmp --to csbu-workmate --apply --json
if ($LASTEXITCODE -ne 0) { throw "theme-remap apply failed`n$apply" }
$applyObj = $apply | ConvertFrom-Json
if ($applyObj.applied -ne $true) { throw 'Apply should set applied=true' }
$saved = Get-Content -LiteralPath $tmp -Raw | ConvertFrom-Json
if ($saved.theme.id -ne 'csbu-workmate') { throw "Expected saved theme csbu-workmate; got $($saved.theme.id)" }
if ($saved.theme.mode -ne 'light') { throw "Expected mode light; got $($saved.theme.mode)" }
if (-not $saved.extensions.themeRemap) { throw 'Expected extensions.themeRemap after apply' }
$validateJson = & dotnet run --project $project -c Release --no-build -- deck validate $tmp --json
$validation = $validateJson | ConvertFrom-Json
$errors = @($validation.diagnostics | Where-Object { $_.severity -eq 'error' })
if ($errors.Count -gt 0) { throw ("Validate errors after theme-remap: {0}" -f (($errors | ForEach-Object code) -join ', ')) }
Write-Output 'apply+validate ok'
Write-Output 'Verify-DeckThemeRemap passed'
