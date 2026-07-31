param(
  [Parameter(Mandatory = $true)]
  [string]$TargetPath
)

$ErrorActionPreference = 'Stop'

$resolvedTarget = (Resolve-Path -LiteralPath $TargetPath).Path
$bytes = [System.IO.File]::ReadAllBytes($resolvedTarget)
$originalLength = $bytes.Length

function Assert-Range {
  param([long]$Offset, [long]$Length, [string]$Description)

  if ($Offset -lt 0 -or $Length -lt 0 -or $Offset + $Length -gt $bytes.Length) {
    throw "Invalid PE $Description range: offset=$Offset length=$Length fileLength=$($bytes.Length)"
  }
}

function Read-UInt16 {
  param([int]$Offset)
  Assert-Range $Offset 2 'UInt16'
  return [System.BitConverter]::ToUInt16($bytes, $Offset)
}

function Read-UInt32 {
  param([int]$Offset)
  Assert-Range $Offset 4 'UInt32'
  return [System.BitConverter]::ToUInt32($bytes, $Offset)
}

function Convert-RvaToOffset {
  param([uint32]$Rva)

  foreach ($section in $sections) {
    $sectionSpan = [Math]::Max([uint64]$section.VirtualSize, [uint64]$section.RawSize)
    if ([uint64]$Rva -ge [uint64]$section.VirtualAddress -and [uint64]$Rva -lt [uint64]$section.VirtualAddress + $sectionSpan) {
      $offset = [uint64]$section.RawPointer + ([uint64]$Rva - [uint64]$section.VirtualAddress)
      Assert-Range $offset 1 'RVA mapping'
      return [int]$offset
    }
  }

  throw ('PE RVA 0x{0:X8} does not map to a file section' -f $Rva)
}

function Get-ResourceEntries {
  param([int]$DirectoryRelativeOffset)

  $directoryOffset = $resourceRootOffset + $DirectoryRelativeOffset
  Assert-Range $directoryOffset 16 'resource directory'
  $entryCount = (Read-UInt16 ($directoryOffset + 12)) + (Read-UInt16 ($directoryOffset + 14))
  $entries = @()

  for ($index = 0; $index -lt $entryCount; $index++) {
    $entryOffset = $directoryOffset + 16 + ($index * 8)
    Assert-Range $entryOffset 8 'resource directory entry'
    $name = Read-UInt32 $entryOffset
    $target = Read-UInt32 ($entryOffset + 4)
    $entries += [pscustomobject]@{
      Id          = [uint32]($name -band 0xFFFF)
      IsNamed     = ($name -band 2147483648) -ne 0
      IsDirectory = ($target -band 2147483648) -ne 0
      Target      = [int]($target -band 0x7FFFFFFF)
    }
  }

  return $entries
}

function Get-ResourceDataEntries {
  param([int]$DirectoryRelativeOffset)

  $results = @()
  foreach ($entry in @(Get-ResourceEntries $DirectoryRelativeOffset)) {
    if ($entry.IsDirectory) {
      $results += @(Get-ResourceDataEntries $entry.Target)
      continue
    }

    $dataEntryOffset = $resourceRootOffset + $entry.Target
    Assert-Range $dataEntryOffset 16 'resource data entry'
    $results += [pscustomobject]@{
      DataRva = Read-UInt32 $dataEntryOffset
      Size    = Read-UInt32 ($dataEntryOffset + 4)
    }
  }

  return $results
}

Assert-Range 0 64 'DOS header'
if ((Read-UInt16 0) -ne 0x5A4D) {
  throw 'Target is not a PE executable: missing MZ header'
}

$peOffset = [int](Read-UInt32 0x3C)
Assert-Range $peOffset 24 'PE header'
if ((Read-UInt32 $peOffset) -ne 0x00004550) {
  throw 'Target is not a PE executable: missing PE signature'
}

$sectionCount = Read-UInt16 ($peOffset + 6)
$optionalHeaderSize = Read-UInt16 ($peOffset + 20)
$optionalHeaderOffset = $peOffset + 24
$optionalHeaderMagic = Read-UInt16 $optionalHeaderOffset
$dataDirectoryOffset = switch ($optionalHeaderMagic) {
  0x10B { $optionalHeaderOffset + 96 }
  0x20B { $optionalHeaderOffset + 112 }
  default { throw ('Unsupported PE optional header magic: 0x{0:X4}' -f $optionalHeaderMagic) }
}

$resourceRva = Read-UInt32 ($dataDirectoryOffset + 16)
$resourceSize = Read-UInt32 ($dataDirectoryOffset + 20)
if ($resourceRva -eq 0 -or $resourceSize -eq 0) {
  Write-Host "No PE resources found in $resolvedTarget"
  exit 0
}

$sectionTableOffset = $optionalHeaderOffset + $optionalHeaderSize
$sections = @()
for ($index = 0; $index -lt $sectionCount; $index++) {
  $sectionOffset = $sectionTableOffset + ($index * 40)
  Assert-Range $sectionOffset 40 'section header'
  $sections += [pscustomobject]@{
    VirtualSize    = Read-UInt32 ($sectionOffset + 8)
    VirtualAddress = Read-UInt32 ($sectionOffset + 12)
    RawSize        = Read-UInt32 ($sectionOffset + 16)
    RawPointer     = Read-UInt32 ($sectionOffset + 20)
  }
}

$resourceRootOffset = Convert-RvaToOffset $resourceRva
$versionType = @(Get-ResourceEntries 0) | Where-Object { -not $_.IsNamed -and $_.Id -eq 16 } | Select-Object -First 1
if (-not $versionType) {
  Write-Host "No VERSIONINFO resource found in $resolvedTarget"
  exit 0
}
if (-not $versionType.IsDirectory) {
  throw 'VERSIONINFO resource entry does not point to a resource directory'
}

$versionDataEntries = @(Get-ResourceDataEntries $versionType.Target)
if ($versionDataEntries.Count -eq 0) {
  throw 'VERSIONINFO resource contains no data entries'
}

foreach ($dataEntry in $versionDataEntries) {
  $dataOffset = Convert-RvaToOffset $dataEntry.DataRva
  Assert-Range $dataOffset $dataEntry.Size 'VERSIONINFO data'
  [System.Array]::Clear($bytes, $dataOffset, [int]$dataEntry.Size)
}

[System.IO.File]::WriteAllBytes($resolvedTarget, $bytes)
if ((Get-Item -LiteralPath $resolvedTarget).Length -ne $originalLength) {
  throw 'In-place VERSIONINFO removal changed the executable length'
}

$versionInfo = (Get-Item -LiteralPath $resolvedTarget).VersionInfo
$remainingValues = @(
  $versionInfo.CompanyName,
  $versionInfo.FileDescription,
  $versionInfo.FileVersion,
  $versionInfo.InternalName,
  $versionInfo.LegalCopyright,
  $versionInfo.LegalTrademarks,
  $versionInfo.OriginalFilename,
  $versionInfo.ProductName,
  $versionInfo.ProductVersion
) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

if ($remainingValues.Count -gt 0) {
  throw "VERSIONINFO removal verification failed: $($remainingValues -join ', ')"
}

Write-Host "Cleared VERSIONINFO in place for $resolvedTarget without changing its length"
