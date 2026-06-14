# COM bridge runner. Run in 32-bit STA PowerShell (VB6 OCXs are x86):
#   powershell.exe (SysWOW64) -STA -File com_bag_decode.ps1 `
#       -OcxPaths <a;b> -ClassName <Class> -LibName <Lib> -TypelibClsids <{..};{..}> `
#       -Versions <2.0;6.0> -EmbeddedClsid <{..}> -BagFile <bag>
# Resolves the control's coclass CLSID from a candidate type library (the .frm
# Object= GUID is the *type library* id, not a coclass): each candidate is loaded
# from its OCX path when present, else from the registry by typelib GUID+version;
# the one whose library/coclass name matches is used. The control is then
# instantiated license-aware, the bag loaded, and gettable properties dumped.
# Emits one line of JSON (see ComBag.cs). Hosts the control in this signed
# powershell.exe process (no unsigned exe spawned, so AV doesn't block it).
# The OcxPaths/TypelibClsids/Versions lists are `;`-separated and index-aligned.
param(
  [string]$OcxPaths = '',
  [string]$ClassName = '',
  [string]$LibName = '',
  [string]$TypelibClsids = '',
  [string]$Versions = '',
  [string]$EmbeddedClsid = '',
  [Parameter(Mandatory = $true)][string]$BagFile
)
$ErrorActionPreference = 'Stop'
try {
  $cs = Get-Content -Raw (Join-Path $PSScriptRoot 'ComBag.cs')
  Add-Type -TypeDefinition $cs -Language CSharp | Out-Null
  $bag = [System.IO.File]::ReadAllBytes($BagFile)
  [ComBag]::Decode($OcxPaths, $ClassName, $LibName, $TypelibClsids, $Versions, $EmbeddedClsid, $bag)
} catch {
  '{"ok":false,"error":' + (ConvertTo-Json $_.Exception.Message) + '}'
}
