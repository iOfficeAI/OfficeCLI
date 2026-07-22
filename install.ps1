$repo = "iOfficeAI/OfficeCLI"
# Select the native asset for the machine architecture. PROCESSOR_ARCHITECTURE
# reports the *current process* arch (x86/AMD64 when an emulated 32/64-bit
# PowerShell runs on ARM64); PROCESSOR_ARCHITEW6432 then reveals the real ARM64
# host, so prefer it when present.
$arch = if ($env:PROCESSOR_ARCHITEW6432) { $env:PROCESSOR_ARCHITEW6432 } else { $env:PROCESSOR_ARCHITECTURE }
$asset = if ($arch -eq "ARM64") { "officecli-win-arm64.exe" } else { "officecli-win-x64.exe" }
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

# Resolve-Version
# Discover the latest release tag (vX.Y.Z) by following the /releases/latest
# redirect and reading the final tag URL. Mirror first, github fallback.
# Returns the tag on success, $null on failure. This lets us download from the
# IMMUTABLE versioned path instead of the mutable /releases/latest/download/
# path — see download section below.
function Resolve-Version {
    foreach ($base in @("$mirrorBase/releases/latest", "https://github.com/$repo/releases/latest")) {
        try {
            $resp = Invoke-WebRequest -Uri $base -MaximumRedirection 5 -TimeoutSec 30 -ErrorAction Stop
            $finalUrl = $resp.BaseResponse.ResponseUri.AbsoluteUri
            if ($finalUrl -match '/releases/tag/(v[0-9]+\.[0-9]+\.[0-9]+)') {
                return $matches[1]
            }
        } catch {
            if ($_.Exception.Response -and $_.Exception.Response.ResponseUri) {
                $finalUrl = $_.Exception.Response.ResponseUri.AbsoluteUri
                if ($finalUrl -match '/releases/tag/(v[0-9]+\.[0-9]+\.[0-9]+)') {
                    return $matches[1]
                }
            }
        }
    }
    return $null
}

$source = $null

# Resolve the latest tag up-front so we download from the IMMUTABLE versioned
# path (/releases/download/vX.Y.Z/asset) instead of the mutable
# /releases/latest/download/ path. The latter is CDN-cached for up to 4h, so
# right after a release it can serve the PREVIOUS binary together with a
# self-consistent stale SHA256SUMS — which passes checksum and installs an old
# version despite printing success. The versioned URL is pinned + immutable.
$version = Resolve-Version
if ($version) {
    Write-Host "Latest version: $version"
    $mirrorAssetBase = "$mirrorBase/releases/download/$version"
    $githubAssetBase = "https://github.com/$repo/releases/download/$version"
} else {
    Write-Host "Could not resolve latest version; falling back to 'latest' path."
    $mirrorAssetBase = "$mirrorBase/releases/latest/download"
    $githubAssetBase = $githubReleaseBase
}

# Step 1: Try downloading (mirror first, github fallback)
$tempFile = "$env:TEMP\$binary"
Write-Host "Downloading OfficeCLI..."
if (Fetch-WithFallback "$mirrorAssetBase/$asset" "$githubAssetBase/$asset" $tempFile) {
    # Verify checksum if available
    $checksumOk = $false
    $checksumFile = "$env:TEMP\officecli-SHA256SUMS"
    if (Fetch-WithFallback "$mirrorAssetBase/SHA256SUMS" "$githubAssetBase/SHA256SUMS" $checksumFile) {
        $checksumContent = Get-Content $checksumFile
        # Match the filename column EXACTLY (not a regex/substring): `-match` is
        # an unanchored regex where `.`/`-` are metacharacters, so it could match
        # the wrong manifest line (or several) and pick a wrong hash — failing an
        # otherwise-valid update. Mirrors the C# self-updater's MatchChecksumManifest.
        $expected = $null
        foreach ($line in $checksumContent) {
            $parts = ($line.Trim() -split '\s+')
            if ($parts.Length -ge 2 -and $parts[1] -eq $asset) { $expected = $parts[0]; break }
        }
        if ($expected) {
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
