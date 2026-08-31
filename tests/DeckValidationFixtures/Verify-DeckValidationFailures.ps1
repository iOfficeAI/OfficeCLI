$ErrorActionPreference = 'Stop'

$project = Join-Path $PSScriptRoot '..\..\src\officecli\officecli.csproj'
$cases = @(
    @{ File = 'missing-slot.workmate-deck.json'; Code = 'block_slot_required' },
    @{ File = 'duplicate-slot.workmate-deck.json'; Code = 'duplicate_slot_assignment' }
)

foreach ($case in $cases) {
    $fixture = Join-Path $PSScriptRoot $case.File
    $output = (& dotnet run --project $project -c Release --no-build -- deck validate $fixture --json 2>&1 | Out-String)
    if ($LASTEXITCODE -eq 0) { throw "Expected deck validation to fail: $($case.File)" }
    $result = $output | ConvertFrom-Json
    if ($result.diagnostics.code -notcontains $case.Code) {
        throw "Expected diagnostic '$($case.Code)' for $($case.File). Output: $output"
    }
}

$outline = Join-Path $PSScriptRoot 'outline-unassigned.workmate-deck.json'
& dotnet run --project $project -c Release --no-build -- deck validate $outline --json
if ($LASTEXITCODE -ne 0) { throw 'Outline decks must allow blocks without assigned layout slots.' }

$previewOutput = (& dotnet run --project $project -c Release --no-build -- deck render $outline --format preview --json | Out-String)
if ($LASTEXITCODE -ne 0) { throw 'Outline preview rendering failed.' }
$preview = $previewOutput | ConvertFrom-Json
$slots = @($preview.slides[0].elements | ForEach-Object slot)
if ($slots.Count -ne 2 -or $slots[0] -ne 'statement' -or $slots[1] -ne 'support') {
    throw "Outline preview reused or misassigned a semantic slot: $($slots -join ', ')"
}

$validDeck = Join-Path $PSScriptRoot '..\DeckFixtures\business-light.workmate-deck.json'
$actualRevision = (Get-Content -LiteralPath $validDeck -Raw | ConvertFrom-Json).revision
$staleOutput = Join-Path ([System.IO.Path]::GetTempPath()) 'officecli-stale-deck-output.pptx'
[System.IO.File]::WriteAllText($staleOutput, 'previous-valid-output')
try {
    $staleResult = (& dotnet run --project $project -c Release --no-build -- deck build $validDeck --output $staleOutput --expected-revision ($actualRevision + 1) --json 2>&1 | Out-String)
    $staleExitCode = $LASTEXITCODE
    if ($staleExitCode -eq 0) { throw 'A stale expected revision must fail the build.' }
    if ([System.IO.File]::ReadAllText($staleOutput) -ne 'previous-valid-output') {
        throw "A stale build replaced the previous output. Command output: $staleResult"
    }
}
finally {
    Remove-Item -LiteralPath $staleOutput -Force -ErrorAction SilentlyContinue
}

# The last native command is expected to fail. Reset its exit code so a successful
# negative-path assertion does not fail PowerShell hosts that propagate LASTEXITCODE.
$global:LASTEXITCODE = 0
