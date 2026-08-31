param(
    [int]$BaselineMilliseconds = 10000,
    [double]$AllowedRegression = 1.2
)

$ErrorActionPreference = 'Stop'

$repositoryRoot = Resolve-Path (Join-Path $PSScriptRoot '..\..')
$project = Join-Path $repositoryRoot 'src\officecli\officecli.csproj'
$fixture = Join-Path $PSScriptRoot 'business-light.workmate-deck.json'
$workingRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('officecli-deck-performance-' + [guid]::NewGuid().ToString('N'))
$resultDirectory = Join-Path $repositoryRoot 'TestResults'
$resultPath = Join-Path $resultDirectory 'deck-performance.json'
[System.IO.Directory]::CreateDirectory($workingRoot) | Out-Null
[System.IO.Directory]::CreateDirectory($resultDirectory) | Out-Null

try {
    $deck = [System.IO.File]::ReadAllText($fixture) | ConvertFrom-Json
    $deck.slides = @($deck.slides | Select-Object -First 10)
    $benchmarkSpec = Join-Path $workingRoot 'benchmark.workmate-deck.json'
    [System.IO.File]::WriteAllText($benchmarkSpec, ($deck | ConvertTo-Json -Depth 100 -Compress))
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'visual.svg') -Destination $workingRoot

    $warmOutput = Join-Path $workingRoot 'warm.pptx'
    & dotnet run --project $project -c Release --no-build -- deck build $benchmarkSpec --output $warmOutput --json | Out-Null
    if ($LASTEXITCODE -ne 0) { throw 'Deck performance warm-up failed.' }

    $samples = 1..3 | ForEach-Object {
        $output = Join-Path $workingRoot ("sample-$_.pptx")
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        & dotnet run --project $project -c Release --no-build -- deck build $benchmarkSpec --output $output --json | Out-Null
        $stopwatch.Stop()
        if ($LASTEXITCODE -ne 0) { throw "Deck performance sample $_ failed." }
        $stopwatch.Elapsed.TotalMilliseconds
    }
    $median = ($samples | Sort-Object)[1]
    $budget = $BaselineMilliseconds * $AllowedRegression
    $result = [ordered]@{
        slides = 10
        samplesMilliseconds = @($samples | ForEach-Object { [Math]::Round($_, 2) })
        medianMilliseconds = [Math]::Round($median, 2)
        baselineMilliseconds = $BaselineMilliseconds
        allowedRegression = $AllowedRegression
        budgetMilliseconds = $budget
        runner = $env:RUNNER_OS
    }
    [System.IO.File]::WriteAllText(
        $resultPath,
        ($result | ConvertTo-Json -Depth 10)
    )
    Write-Output ("10-slide median: {0:N2} ms (budget: {1:N0} ms)" -f $median, $budget)
    if ($median -gt $budget) {
        throw ("Deck compilation regressed beyond the 20% budget: {0:N2} ms > {1:N0} ms." -f $median, $budget)
    }
}
finally {
    if ([System.IO.Directory]::Exists($workingRoot)) {
        [System.IO.Directory]::Delete($workingRoot, $true)
    }
}
