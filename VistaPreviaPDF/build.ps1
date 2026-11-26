$ErrorActionPreference = 'Stop'
$csc = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
$fw  = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319"
$wm  = "C:\Windows\System32\WinMetadata"
$gac = "C:\Windows\Microsoft.NET\assembly\GAC_MSIL"
$dir = Split-Path -Parent $MyInvocation.MyCommand.Path

$uiaClient = (Get-ChildItem "$gac\UIAutomationClient" -Recurse -Filter UIAutomationClient.dll | Select-Object -First 1).FullName
$uiaTypes  = (Get-ChildItem "$gac\UIAutomationTypes" -Recurse -Filter UIAutomationTypes.dll | Select-Object -First 1).FullName
$winBase   = (Get-ChildItem "$gac\WindowsBase" -Recurse -Filter WindowsBase.dll | Select-Object -First 1).FullName
$winRt     = (Get-ChildItem "$gac\System.Runtime.WindowsRuntime" -Recurse -Filter System.Runtime.WindowsRuntime.dll | Select-Object -First 1).FullName

& $csc /nologo /target:winexe /out:"$dir\VistaPreviaPDF.exe" /win32icon:"$dir\app.ico" `
  /reference:"$fw\System.dll" `
  /reference:"$fw\System.Core.dll" `
  /reference:"$fw\System.Drawing.dll" `
  /reference:"$fw\System.Windows.Forms.dll" `
  /reference:"$fw\Microsoft.CSharp.dll" `
  /reference:"$uiaClient" `
  /reference:"$uiaTypes" `
  /reference:"$winBase" `
  /reference:"$fw\System.Runtime.dll" `
  /reference:"$winRt" `
  /reference:"$wm\Windows.Foundation.winmd" `
  /reference:"$wm\Windows.Storage.winmd" `
  /reference:"$wm\Windows.Data.winmd" `
  "$dir\PdfHoverPreview.cs"

Write-Host "exe existe: $(Test-Path "$dir\VistaPreviaPDF.exe")"
