param([string]$Path)
$clsid = 'HKCU:\Software\Classes\CLSID\{000209FF-0000-0000-C000-000000000046}\LocalServer32'
$old = $null; if (Test-Path $clsid) { $old = (Get-ItemProperty $clsid).'(default)' }
if ($old) { Remove-Item -Path $clsid -Force }
$full = [System.IO.Path]::GetFullPath($Path)
Remove-Item -Force "$env:APPDATA\Microsoft\Word\*.asd" -ErrorAction SilentlyContinue
Remove-Item -Force ([System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($full), "~`$" + [System.IO.Path]::GetFileName($full))) -ErrorAction SilentlyContinue
$wd = $null
try {
  # Word remembers every path it failed to open in `Resiliency\DisabledItems` and
  # then hangs on that path forever, even after the file is fixed. Clear the
  # list before each run so a repaired sample can be checked again.
  Remove-Item -Path 'HKCU:\Software\Microsoft\Office\16.0\Word\Resiliency' -Recurse -Force -ErrorAction SilentlyContinue
  $wd = New-Object -ComObject Word.Application
  $wd.Visible = $false
  $wd.DisplayAlerts = 0
  Write-Output ("app " + $wd.Path)
  Write-Output "step opening"
  $doc = $wd.Documents.Open($full, $false, $true)
  Write-Output ("OPEN paras=" + $doc.Paragraphs.Count + " f=" + $doc.Footnotes.Count + " e=" + $doc.Endnotes.Count)
  $doc.Close($false)
} catch {
  Write-Output ("ERR line " + $_.InvocationInfo.ScriptLineNumber + ": " + $_.Exception.Message)
} finally {
  if ($wd) { try { $wd.Quit() } catch {} }
  Start-Sleep -Milliseconds 300
  Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
  if ($old) { New-Item -Path $clsid -Force | Out-Null; Set-ItemProperty -Path $clsid -Name '(default)' -Value $old }
}