param([string]$Out = "samples/ms/footnote2.doc")
# One footnote, no header: the smallest document that still has the two PLCs,
# so their "last CP" convention can be read without a header story to confuse it.
$ErrorActionPreference = 'Continue'
$clsid = 'HKCU:\Software\Classes\CLSID\{000209FF-0000-0000-C000-000000000046}\LocalServer32'
$old = $null
if (Test-Path $clsid) { $old = (Get-ItemProperty $clsid).'(default)' }
if ($old) { Remove-Item -Path $clsid -Force }
Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
try {
  $app = New-Object -ComObject Word.Application
  $app.Visible = $false
  $app.DisplayAlerts = 0
  Write-Output ("app path: " + $app.Path)
  $doc = $app.Documents.Add()
  $body = "Alpha beta gamma.`rSecond paragraph."
  $doc.Content.Text = $body
  $at = $body.IndexOf("gamma.") + 6
  $doc.Range($at, $at).Footnotes.Add($doc.Range($at, $at), [System.Type]::Missing, "one note") | Out-Null
  Write-Output ("note at " + $at + ", count=" + $doc.Footnotes.Count)
  $path = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot "..\$Out"))
  $doc.SaveAs2([string]$path, 0)
  Write-Output ("saved " + $path)
  $doc.Close($false)
  $app.Quit()
} catch {
  Write-Output ("ERR " + $_.Exception.Message + " | line " + $_.InvocationInfo.ScriptLineNumber)
} finally {
  Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
  if ($old) { New-Item -Path $clsid -Force | Out-Null; Set-ItemProperty -Path $clsid -Name '(default)' -Value $old }
}