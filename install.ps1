$repo = "iOfficeAI/OfficeCLI"
$asset = "officecli-win-x64.exe"
$binary = "officecli.exe"

# Mirror primary, github fallback. The mirror is exercised first so issues
# surface there fast; github is the safety net when CF or the mirror is
# unreachable.
$mirrorBase = "https://d.officecli.ai"
$githubReleaseBase = "https://github.com/$repo/releases/latest/download"
$githubRawBase = "https://raw.githubusercontent.com/$repo/main"

function Fetch-WithFallback {
    param([string]$Primary, [string]$Fallback, [string]$OutFile)
    try {
        Invoke-WebRequest -Uri $Primary -OutFile $OutFile -TimeoutSec 30 -ErrorAction Stop
        Write-Host "  (via mirror)"
        return $true
    } catch {
        Write-Host "  mirror unreachable, falling back to github..."
        try {
            Invoke-WebRequest -Uri $Fallback -OutFile $OutFile -TimeoutSec 300 -ErrorAction Stop
            return $true
        } catch {
            return $false
        }
    }
}

$source = $null

# Step 1: Try downloading (mirror first, github fallback)
$tempFile = "$env:TEMP\$binary"
$assetBase = $asset -replace '\.exe$', ''
Write-Host "Downloading OfficeCLI..."
if (Fetch-WithFallback "$mirrorBase/releases/latest/download/$asset" "$githubReleaseBase/$asset" $tempFile) {
    # Verify checksum if available
    $checksumOk = $false
    $checksumFile = "$env:TEMP\officecli-SHA256SUMS"
    if (Fetch-WithFallback "$mirrorBase/releases/latest/download/SHA256SUMS" "$githubReleaseBase/SHA256SUMS" $checksumFile) {
        $checksumContent = Get-Content $checksumFile
        $expectedLine = $checksumContent | Where-Object { $_ -match $asset }
        if ($expectedLine) {
            $expected = ($expectedLine -split '\s+')[0]
            $actual = (Get-FileHash -Path $tempFile -Algorithm SHA256).Hash.ToLower()
            if ($expected -eq $actual) {
                $checksumOk = $true
                Write-Host "Checksum verified."
            } else {
                Write-Host "Checksum mismatch! Expected: $expected, Got: $actual"
                Remove-Item -Force $tempFile, $checksumFile -ErrorAction SilentlyContinue
                exit 1
            }
        }
        Remove-Item -Force $checksumFile -ErrorAction SilentlyContinue
    } else {
        Write-Host "Checksum file not available, skipping verification."
    }
    $output = & $tempFile --version 2>&1
    if ($LASTEXITCODE -eq 0) {
        $source = $tempFile
        Write-Host "Download verified."
    } else {
        Write-Host "Downloaded file is not a valid OfficeCLI binary."
        Remove-Item -Force $tempFile -ErrorAction SilentlyContinue
    }
} else {
    Write-Host "Download failed."
}

# Step 2: Fallback to local files
if (-not $source) {
    Write-Host "Looking for local binary..."
    $candidates = @(".\$asset", ".\$binary", ".\bin\$asset", ".\bin\$binary", ".\bin\release\$asset", ".\bin\release\$binary")
    foreach ($candidate in $candidates) {
        if (Test-Path $candidate) {
            $output = & $candidate --version 2>&1
            if ($LASTEXITCODE -eq 0) {
                $source = $candidate
                Write-Host "Found valid binary at $candidate"
                break
            }
        }
    }
}

if (-not $source) {
    Write-Host "Error: Could not find a valid OfficeCLI binary."
    Write-Host "Download manually from: https://github.com/$repo/releases"
    exit 1
}

# Step 3: Install
$existing = Get-Command $binary -ErrorAction SilentlyContinue
if ($existing) {
    $installDir = Split-Path $existing.Source
    Write-Host "Found existing installation at $($existing.Source), upgrading..."
} else {
    $installDir = "$env:LOCALAPPDATA\OfficeCLI"
}

New-Item -ItemType Directory -Force -Path $installDir | Out-Null
Copy-Item -Force $source "$installDir\$binary"

foreach ($sidecar in @("rhwp-field-bridge", "rhwp-officecli-bridge")) {
    $sidecarAsset = "$assetBase-$sidecar.exe"
    $sidecarTemp = "$env:TEMP\$sidecarAsset"
    $sidecarTarget = Join-Path $installDir "$sidecar.exe"
    $sidecarSource = $null

    Write-Host "Checking optional HWP sidecar $sidecarAsset..."
    try {
        Invoke-WebRequest -Uri "https://github.com/$repo/releases/latest/download/$sidecarAsset" -OutFile $sidecarTemp
        $sidecarSource = $sidecarTemp
    } catch {
        $candidates = @(".\$sidecarAsset", ".\bin\$sidecarAsset", ".\bin\release\$sidecarAsset", ".\$sidecar.exe", ".\bin\$sidecar.exe", ".\bin\release\$sidecar.exe")
        foreach ($candidate in $candidates) {
            if (Test-Path $candidate) {
                $sidecarSource = $candidate
                break
            }
        }
    }

    if ($sidecarSource) {
        Copy-Item -Force $sidecarSource $sidecarTarget
        Write-Host "Installed HWP sidecar: $sidecarTarget"
    } else {
        Write-Host "Optional HWP sidecar unavailable: $sidecarAsset. Binary .hwp create/read/edit will be dependency-gated."
    }
    Remove-Item -Force $sidecarTemp -ErrorAction SilentlyContinue
}

Remove-Item -Force $tempFile -ErrorAction SilentlyContinue

# Add to PATH if not already there
$currentPath = [Environment]::GetEnvironmentVariable("Path", "User")
if ($currentPath -notlike "*$installDir*") {
    [Environment]::SetEnvironmentVariable("Path", "$currentPath;$installDir", "User")
    Write-Host "Added $installDir to PATH (restart your terminal to take effect)."
}

# Step 4: Install AI agent skills (first install only)
$skillMarker = "$installDir\.officecli-skills-installed"
if (-not (Test-Path $skillMarker)) {
    $skillTargets = @()
    $tools = @{
        "$env:USERPROFILE\.claude" = "Claude Code"
        "$env:USERPROFILE\.copilot" = "GitHub Copilot"
        "$env:USERPROFILE\.agents" = "Codex CLI"
        "$env:USERPROFILE\.cursor" = "Cursor"
        "$env:USERPROFILE\.windsurf" = "Windsurf"
        "$env:USERPROFILE\.minimax" = "MiniMax CLI"
        "$env:USERPROFILE\.openclaw" = "OpenClaw"
        "$env:USERPROFILE\.nanobot\workspace" = "NanoBot"
        "$env:USERPROFILE\.zeroclaw\workspace" = "ZeroClaw"
        "$env:USERPROFILE\.hermes" = "Hermes Agent"
    }
    foreach ($dir in $tools.Keys) {
        if (Test-Path $dir) {
            $skillTargets += "$dir\skills\officecli"
            Write-Host "$($tools[$dir]) detected."
        }
    }

    if ($skillTargets.Count -gt 0) {
        Write-Host "Downloading officecli skill..."
        $tempSkill = "$env:TEMP\officecli-skill.md"
        if (Fetch-WithFallback "$mirrorBase/SKILL.md" "$githubRawBase/SKILL.md" $tempSkill) {
            foreach ($target in $skillTargets) {
                New-Item -ItemType Directory -Force -Path $target | Out-Null
                Copy-Item -Force $tempSkill "$target\SKILL.md"
                Write-Host "  Installed: $target\SKILL.md"
            }
            Remove-Item -Force $tempSkill -ErrorAction SilentlyContinue
        }
    }
    New-Item -ItemType File -Force -Path $skillMarker | Out-Null
}

Write-Host "OfficeCLI installed successfully!"
Write-Host "Run 'officecli --help' to get started."
