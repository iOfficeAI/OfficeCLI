$repo = "iOfficeAI/OfficeCli"
$asset = "officecli-win-x64.exe"
$binary = "officecli.exe"

$source = $null

# Step 1: Try downloading from GitHub
$url = "https://github.com/$repo/releases/latest/download/$asset"
$tempFile = "$env:TEMP\$binary"
Write-Host "Downloading OfficeCli..."
try {
    Invoke-WebRequest -Uri $url -OutFile $tempFile
    $output = & $tempFile --version 2>&1
    if ($LASTEXITCODE -eq 0) {
        $source = $tempFile
        Write-Host "Download verified."
    } else {
        Write-Host "Downloaded file is not a valid OfficeCli binary."
        Remove-Item -Force $tempFile -ErrorAction SilentlyContinue
    }
} catch {
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
    Write-Host "Error: Could not find a valid OfficeCli binary."
    Write-Host "Download manually from: https://github.com/$repo/releases"
    exit 1
}

# Step 3: Install
$existing = Get-Command $binary -ErrorAction SilentlyContinue
if ($existing) {
    $installDir = Split-Path $existing.Source
    Write-Host "Found existing installation at $($existing.Source), upgrading..."
} else {
    $installDir = "$env:LOCALAPPDATA\OfficeCli"
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

# Step 4: Install AI agent skills for detected tools
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
    try {
        Invoke-WebRequest -Uri "https://raw.githubusercontent.com/$repo/main/SKILL.md" -OutFile $tempSkill
        foreach ($target in $skillTargets) {
            New-Item -ItemType Directory -Force -Path $target | Out-Null
            Copy-Item -Force $tempSkill "$target\SKILL.md"
            Write-Host "  Installed: $target\SKILL.md"
        }
        Remove-Item -Force $tempSkill -ErrorAction SilentlyContinue
    } catch {}
}

Write-Host "OfficeCli installed successfully!"
Write-Host "Run 'officecli --help' to get started."
