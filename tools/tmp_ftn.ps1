param([string]$Out = "samples/ms/footnote.doc")
# Word 2016 sample with two footnotes, marks placed at explicit character
# offsets in an already-typed body. Adding notes through the live Selection put
# the marks at the very end of the document, so the ranges are addressed
# directly -- and the later offset is used first so the earlier one stays valid.
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
  $body = "Body one comes before the mark. and continues after it.`rBody two is short.`rBody three has no note."
  $doc.Content.Text = $body
  Write-Output ("body chars=" + $doc.Content.Characters.Count)
  $second = $body.IndexOf("short.") + 6
  $first = $body.IndexOf("mark.") + 5
  $doc.Range($second, $second).Footnotes.Add($doc.Range($second, $second), [System.Type]::Missing, "second note, longer than the first") | Out-Null
  Write-Output ("note at $second, count=" + $doc.Footnotes.Count)
  $doc.Range($first, $first).Footnotes.Add($doc.Range($first, $first), [System.Type]::Missing, "first note") | Out-Null
  Write-Output ("note at $first, count=" + $doc.Footnotes.Count)
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