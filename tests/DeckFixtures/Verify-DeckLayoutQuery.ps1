$ErrorActionPreference = 'Stop'

$project = Join-Path $PSScriptRoot '..\..\src\officecli\officecli.csproj'

function Invoke-LayoutQuery([string[]]$Args) {
    $json = & dotnet run --project $project -c Release --no-build -- deck layout-query @Args --json
    if ($LASTEXITCODE -ne 0) { throw "layout-query failed: $($Args -join ' ')`n$json" }
    return ($json | ConvertFrom-Json)
}

$metrics = Invoke-LayoutQuery @('--role', 'metrics', '--item-count', '4', '--limit', '5')
if ($metrics.resultCount -lt 1) { throw 'Expected metrics results' }
if ($metrics.results[0].role -ne 'metrics') { throw "Top hit role was $($metrics.results[0].role)" }
$topIds = @($metrics.results | ForEach-Object layoutId)
if ($topIds -notcontains 'metrics-row-4' -and $topIds -notcontains 'metrics-strip' -and $topIds -notcontains 'kpi-sparkline-row') {
    throw "Expected a 4-wide metrics layout near the top; got: $($topIds -join ', ')"
}
Write-Output ("metrics top: {0} score={1} capacity={2}" -f $metrics.results[0].layoutId, $metrics.results[0].score, $metrics.results[0].capacity)

$chart = Invoke-LayoutQuery @('--role', 'trend', '--has-chart', 'true', '--item-count', '2', '--limit', '5')
$chartAccepts = @($chart.results[0].accepts)
if ($chartAccepts -notcontains 'chart') { throw "Top trend hit should accept chart; got $($chart.results[0].layoutId)" }
Write-Output ("trend/chart top: {0}" -f $chart.results[0].layoutId)

$media = Invoke-LayoutQuery @('--role', 'image', '--needs-media', 'true', '--item-count', '3', '--limit', '5')
$mediaAccepts = @($media.results[0].accepts)
if ($mediaAccepts -notcontains 'image') { throw "Top image hit should accept image; got $($media.results[0].layoutId)" }
Write-Output ("image/media top: {0}" -f $media.results[0].layoutId)

$query = Invoke-LayoutQuery @('--query', 'waterfall', '--limit', '3')
$ids = @($query.results | ForEach-Object layoutId)
if ($ids -notcontains 'chart-waterfall') { throw "Expected chart-waterfall in query=waterfall results; got $($ids -join ', ')" }
Write-Output ("query waterfall: {0}" -f ($ids -join ', '))

# candidates[] validation fixture (unknown id must fail validate)
$tmp = Join-Path ([System.IO.Path]::GetTempPath()) 'officecli-layout-query-candidates.workmate-deck.json'
@'
{
  "schemaVersion": 1,
  "revision": 1,
  "stage": "outline",
  "metadata": { "title": "Candidates probe", "language": "en-US", "aspectRatio": "16:9" },
  "theme": { "id": "business-light" },
  "slides": [
    {
      "id": "s1",
      "role": "metrics",
      "layoutId": "metrics",
      "candidates": ["metrics-row-4", "not-a-real-layout"],
      "blocks": []
    }
  ],
  "assets": []
}
'@ | Set-Content -LiteralPath $tmp -Encoding utf8
$validateJson = & dotnet run --project $project -c Release --no-build -- deck validate $tmp --json
if ($LASTEXITCODE -eq 0) { throw 'Expected validate to fail for unknown layout candidate' }
$validation = $validateJson | ConvertFrom-Json
$codes = @($validation.diagnostics | ForEach-Object code)
if ($codes -notcontains 'unknown_layout_candidate') { throw "Expected unknown_layout_candidate; got $($codes -join ', ')" }
Write-Output 'candidates validation: unknown_layout_candidate OK'

# valid candidates should pass
@'
{
  "schemaVersion": 1,
  "revision": 1,
  "stage": "outline",
  "metadata": { "title": "Candidates ok", "language": "en-US", "aspectRatio": "16:9" },
  "theme": { "id": "business-light" },
  "slides": [
    {
      "id": "s1",
      "role": "metrics",
      "layoutId": "metrics",
      "candidates": ["metrics-row-4", "kpi-trio"],
      "blocks": []
    }
  ],
  "assets": []
}
'@ | Set-Content -LiteralPath $tmp -Encoding utf8
& dotnet run --project $project -c Release --no-build -- deck validate $tmp --json | Out-Null
if ($LASTEXITCODE -ne 0) { throw 'Expected validate to pass for known layout candidates' }
Write-Output 'candidates validation: known ids OK'
Write-Output 'Verify-DeckLayoutQuery passed'
