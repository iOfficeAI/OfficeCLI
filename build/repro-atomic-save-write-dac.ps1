param([Parameter(Mandatory)][string]$OfficeCli)
$ErrorActionPreference = 'Stop'
$OfficeCli = (Resolve-Path -LiteralPath $OfficeCli).Path
$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if ($principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { throw 'Run this script in a normal PowerShell window, without Run as administrator.' }
$env:OFFICECLI_NO_AUTO_INSTALL = '1'
$env:OFFICECLI_SKIP_UPDATE = '1'
$env:OFFICECLI_NO_AUTO_RESIDENT = '1'
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class OfficeAtomicAcl {
  [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
  public static extern IntPtr CreateFileW(string path, uint access, uint share, IntPtr security, uint creation, uint flags, IntPtr template);
  [DllImport("advapi32.dll", SetLastError=true)]
  public static extern uint GetSecurityInfo(IntPtr handle, uint type, uint info, IntPtr owner, IntPtr group, out IntPtr dacl, IntPtr sacl, out IntPtr descriptor);
  [DllImport("advapi32.dll", SetLastError=true)]
  public static extern uint SetSecurityInfo(IntPtr handle, uint type, uint info, IntPtr owner, IntPtr group, IntPtr dacl, IntPtr sacl);
  [DllImport("kernel32.dll")] public static extern bool CloseHandle(IntPtr handle);
  [DllImport("kernel32.dll")] public static extern IntPtr LocalFree(IntPtr memory);
}
'@
$root = Join-Path ([IO.Path]::GetTempPath()) ('officecli-atomic-save-' + [guid]::NewGuid().ToString('n'))
New-Item -ItemType Directory -Path $root | Out-Null
$handle = [OfficeAtomicAcl]::CreateFileW($root, 0x60000, 7, [IntPtr]::Zero, 3, 0x02000000, [IntPtr]::Zero)
if ($handle -eq [IntPtr](-1)) { throw 'Cannot retain the temporary folder permissions for cleanup.' }
$dacl = [IntPtr]::Zero; $descriptor = [IntPtr]::Zero
$read = [OfficeAtomicAcl]::GetSecurityInfo($handle, 1, 4, [IntPtr]::Zero, [IntPtr]::Zero, [ref]$dacl, [IntPtr]::Zero, [ref]$descriptor)
if ($read -ne 0) { [void][OfficeAtomicAcl]::CloseHandle($handle); throw 'Cannot read the temporary folder permissions.' }
try {
  $sid = [Security.Principal.WindowsIdentity]::GetCurrent().User
  $ownerRights = New-Object Security.Principal.SecurityIdentifier('S-1-3-4')
  $inherit = [Security.AccessControl.InheritanceFlags]'ContainerInherit,ObjectInherit'
  $acl = Get-Acl -LiteralPath $root
  # OWNER RIGHTS suppresses the owner's implicit Change permissions grant.
  # The explicit deny applies to both this folder and its new files; ordinary
  # read/write/delete rights remain as granted by the original directory ACL.
  $acl.AddAccessRule((New-Object Security.AccessControl.FileSystemAccessRule($ownerRights, [Security.AccessControl.FileSystemRights]::ReadPermissions, $inherit, [Security.AccessControl.PropagationFlags]::None, [Security.AccessControl.AccessControlType]::Allow)))
  $acl.AddAccessRule((New-Object Security.AccessControl.FileSystemAccessRule($sid, [Security.AccessControl.FileSystemRights]::ChangePermissions, $inherit, [Security.AccessControl.PropagationFlags]::None, [Security.AccessControl.AccessControlType]::Deny)))
  Set-Acl -LiteralPath $root -AclObject $acl
  $probe = [OfficeAtomicAcl]::CreateFileW($root, 0x40000, 7, [IntPtr]::Zero, 3, 0x02000000, [IntPtr]::Zero)
  $denied = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
  if ($probe -ne [IntPtr](-1)) { [void][OfficeAtomicAcl]::CloseHandle($probe); throw 'The test folder still allows Change permissions.' }
  if ($denied -ne 5) { throw 'The test folder did not reproduce access denied for Change permissions.' }
  Write-Host 'Temporary folder denies Change permissions; ordinary document writes remain allowed.'
  $doc = Join-Path $root 'atomic-save.docx'
  & $OfficeCli create $doc
  if ($LASTEXITCODE -ne 0) { throw 'Creating the document failed.' }
  & $OfficeCli add $doc /body --type paragraph --prop 'text=Saved without Change permissions'
  $editExit = $LASTEXITCODE
  $zip = [IO.Compression.ZipFile]::OpenRead($doc)
  try {
    $reader = New-Object IO.StreamReader($zip.GetEntry('word/document.xml').Open())
    try { $text = $reader.ReadToEnd() } finally { $reader.Dispose() }
  } finally { $zip.Dispose() }
  if ($text -notmatch 'Saved without Change permissions') {
    throw "OfficeCLI edit exit=$editExit, but the saved document still lacks the requested paragraph."
  }
  if ($editExit -ne 0) { throw 'The saved edit reported a failure.' }
  $input = Join-Path $root 'batch.json'
  '[{"command":"add","parent":"/body","type":"paragraph","props":{"text":"Saved through atomic batch"}}]' | Set-Content -LiteralPath $input -Encoding utf8
  & $OfficeCli batch $doc --input $input
  if ($LASTEXITCODE -ne 0) { throw 'The atomic batch failed.' }
  $zip = [IO.Compression.ZipFile]::OpenRead($doc)
  try {
    $reader = New-Object IO.StreamReader($zip.GetEntry('word/document.xml').Open())
    try { $text = $reader.ReadToEnd() } finally { $reader.Dispose() }
  } finally { $zip.Dispose() }
  if ($text -notmatch 'Saved through atomic batch') { throw 'The atomic batch did not persist its paragraph.' }
  Write-Host 'PASS: standalone edit and atomic batch persisted the requested paragraphs.'
} finally {
  # The handle was opened before the deny. Restore the original DACL through
  # that handle, which also removes the inherited deny from the temporary files.
  $restored = [OfficeAtomicAcl]::SetSecurityInfo($handle, 1, 4, [IntPtr]::Zero, [IntPtr]::Zero, $dacl, [IntPtr]::Zero)
  [void][OfficeAtomicAcl]::CloseHandle($handle)
  [void][OfficeAtomicAcl]::LocalFree($descriptor)
  if ($restored -eq 0) { Remove-Item -LiteralPath $root -Recurse -Force }
  else { Write-Warning 'Could not restore the temporary folder permissions; the diagnostic folder was kept.' }
}
