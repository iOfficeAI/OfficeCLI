$ErrorActionPreference = 'Stop'

$project = Join-Path $PSScriptRoot '..\..\src\officecli\officecli.csproj'
$outputRoot = Join-Path ([System.IO.Path]::GetTempPath()) 'officecli-deck-fixtures'
[System.IO.Directory]::CreateDirectory($outputRoot) | Out-Null

Get-ChildItem -LiteralPath $PSScriptRoot -Filter '*.workmate-deck.json' | ForEach-Object {
    $fixture = $_
    & dotnet run --project $project -c Release --no-build -- deck validate $fixture.FullName --json
    if ($LASTEXITCODE -ne 0) { throw "Deck validation failed: $($fixture.Name)" }

    $pptx = Join-Path $outputRoot ($fixture.BaseName + '.pptx')
    $revision = (Get-Content -LiteralPath $fixture.FullName -Raw | ConvertFrom-Json).revision
    & dotnet run --project $project -c Release --no-build -- deck build $fixture.FullName --output $pptx --expected-revision $revision --json
    if ($LASTEXITCODE -ne 0) { throw "Deck build failed: $($fixture.Name)" }

    $archive = [System.IO.Compression.ZipFile]::OpenRead($pptx)
    try {
        $names = @($archive.Entries | ForEach-Object FullName)
        $slideCount = @($names | Where-Object { $_ -match '^ppt/slides/slide\d+\.xml$' }).Count
        $notesCount = @($names | Where-Object { $_ -match '^ppt/notesSlides/notesSlide\d+\.xml$' }).Count
        $chartEntries = @($names | Where-Object { $_ -match '^ppt/(?:[^/]+/)*charts/chart(?:Ex)?\d*\.xml$' })
        $chartCount = $chartEntries.Count
        $tableCount = 0
        $archive.Entries | Where-Object { $_.FullName -match '^ppt/slides/slide\d+\.xml$' } | ForEach-Object {
            $reader = [System.IO.StreamReader]::new($_.Open())
            try { if ($reader.ReadToEnd().Contains('<a:tbl>')) { $tableCount += 1 } }
            finally { $reader.Dispose() }
        }
        if ($slideCount -ne 12) { throw "Expected 12 slides in $($fixture.Name), found $slideCount" }
        if ($notesCount -ne 12) { throw "Expected editable notes on all 12 slides in $($fixture.Name), found $notesCount" }
        if ($chartCount -lt 1) {
            $actualChartEntries = @($names | Where-Object { $_ -match '^ppt/(?:[^/]+/)*charts/' }) -join ', '
            throw "Expected an editable chart in $($fixture.Name); chart entries: $actualChartEntries"
        }
        if ($tableCount -lt 1) { throw "Expected an editable table in $($fixture.Name)" }
    }
    finally {
        $archive.Dispose()
    }
}
