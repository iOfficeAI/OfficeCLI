$ErrorActionPreference = 'Stop'

$project = Join-Path $PSScriptRoot '..\..\src\officecli\officecli.csproj'

$catalogJson = & dotnet run --project $project -c Release --no-build -- deck catalog --json
if ($LASTEXITCODE -ne 0) { throw "catalog failed`n$catalogJson" }
$catalog = $catalogJson | ConvertFrom-Json
if ($catalog.version -ne '1.5.3') { throw "Expected catalog 1.5.3; got $($catalog.version)" }
$controlIds = @{}
foreach ($layout in @($catalog.layouts)) {
    foreach ($ctl in @($layout.controls)) { if ($ctl.id) { $controlIds[$ctl.id] = $true } }
}
foreach ($need in @('showKicker', 'focusIndex')) {
    if (-not $controlIds.ContainsKey($need)) { throw "Expected control key $need in catalog" }
}
Write-Output ("catalog ok: version={0} controlKeys={1}" -f $catalog.version, ($controlIds.Keys.Count))

$tmpDir = Join-Path ([System.IO.Path]::GetTempPath()) 'officecli-deck-scaffold'
[System.IO.Directory]::CreateDirectory($tmpDir) | Out-Null
$out = Join-Path $tmpDir 'long-outline.workmate-deck.json'

$json = & dotnet run --project $project -c Release --no-build -- deck scaffold `
    --goal "Board update: revenue, pipeline, risks" `
    --audience "board directors" `
    --pages 20 `
    --theme csbu-workmate `
    --seed board-2026 `
    -o $out `
    --json
if ($LASTEXITCODE -ne 0) { throw "scaffold failed`n$json" }
$obj = $json | ConvertFrom-Json
if ($obj.written -ne $true) { throw 'Expected written=true' }
if ($obj.slideCount -ne 20) { throw "Expected 20 slides; got $($obj.slideCount)" }
if ($obj.stage -ne 'outline') { throw "Expected stage outline; got $($obj.stage)" }
if ($obj.report.seed -ne 'board-2026') { throw "Expected seed board-2026; got $($obj.report.seed)" }
if ($obj.report.sectionBreaks -lt 1) { throw 'Expected at least one section transition on 20-page deck' }

$spec = Get-Content -LiteralPath $out -Raw | ConvertFrom-Json
if ($spec.stage -ne 'outline') { throw 'Saved spec stage mismatch' }
if ($spec.theme.id -ne 'csbu-workmate') { throw "Expected theme csbu-workmate; got $($spec.theme.id)" }
if (-not $spec.extensions.deckScaffold) { throw 'Expected extensions.deckScaffold' }
if (@($spec.slides).Count -ne 20) { throw 'Saved slide count mismatch' }
$roles = @($spec.slides | ForEach-Object role)
if ($roles[0] -ne 'cover') { throw 'First slide should be cover' }
if ($roles[-1] -ne 'closing') { throw 'Last slide should be closing' }
if ($roles -notcontains 'transition') { throw 'Expected a transition role in long scaffold' }

# Same seed → same role sequence
$out2 = Join-Path $tmpDir 'long-outline-b.workmate-deck.json'
$json2 = & dotnet run --project $project -c Release --no-build -- deck scaffold `
    --goal "Board update: revenue, pipeline, risks" `
    --audience "board directors" `
    --pages 20 `
    --theme csbu-workmate `
    --seed board-2026 `
    -o $out2 `
    --json
if ($LASTEXITCODE -ne 0) { throw "scaffold rerun failed`n$json2" }
$spec2 = Get-Content -LiteralPath $out2 -Raw | ConvertFrom-Json
$roles2 = @($spec2.slides | ForEach-Object role)
if (($roles -join ',') -ne ($roles2 -join ',')) { throw 'Seeded scaffolds should match role sequence' }

$validateJson = & dotnet run --project $project -c Release --no-build -- deck validate $out --json
if ($LASTEXITCODE -ne 0) { throw "validate failed`n$validateJson" }
$validation = $validateJson | ConvertFrom-Json
$errors = @($validation.diagnostics | Where-Object { $_.severity -eq 'error' })
if ($errors.Count -gt 0) { throw ("Validate errors after scaffold: {0}" -f (($errors | ForEach-Object code) -join ', ')) }

Write-Output 'Verify-DeckScaffold passed'
