# Tier-3 COM bridge runner. Run in 32-bit STA PowerShell (VB6 OCXs are x86):
#   powershell.exe (SysWOW64) -STA -File com_bag_decode.ps1 <CLSID> <bagfile>
# Emits one line of JSON (see ComBag.cs). Hosts the control in this signed
# powershell.exe process (no unsigned exe spawned, so AV doesn't block it).
param(
  [Parameter(Mandatory = $true)][string]$Clsid,
  [Parameter(Mandatory = $true)][string]$BagFile
)
$ErrorActionPreference = 'Stop'
try {
  $cs = Get-Content -Raw (Join-Path $PSScriptRoot 'ComBag.cs')
  Add-Type -TypeDefinition $cs -Language CSharp | Out-Null
  $bag = [System.IO.File]::ReadAllBytes($BagFile)
  [ComBag]::Decode($Clsid, $bag)
} catch {
  '{"ok":false,"error":' + (ConvertTo-Json $_.Exception.Message) + '}'
}
