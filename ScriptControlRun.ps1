param ([Parameter(Mandatory=$true)][string]$File)

if ($env:PROCESSOR_ARCHITECTURE -ne 'x86') {
  throw "MSScriptControl.ScriptControl only in x86 mode"
}
if (!(Test-Path $File -PathType Leaf)) {
  throw (New-Object System.IO.FileNotFoundException $File)
}

$sc = New-Object -ComObject MSScriptControl.ScriptControl
$sc.Language = "VBS"

$sc.ExecuteStatement((Get-Content $file))