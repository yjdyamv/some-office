param([string]$Path = "samples/ms/bookmarks.doc")
# Build a small document carrying bookmarks, using real Word.
$ErrorActionPreference = 'Continue'
$clsid = 'HKCU:\Software\Classes\CLSID\{000209FF-0000-0000-C000-000000000046}\LocalServer32'
$old = $null
if (Test-Path $clsid) { $old = (Get-ItemProperty $clsid).'(default)' }
if ($old) { Remove-Item -Path $clsid -Force }
Remove-Item -Path 'HKCU:\Software\Microsoft\Office\16.0\Word\Resiliency' -Recurse -Force -ErrorAction SilentlyContinue
Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
try {
  $wd = New-Object -ComObject Word.Application
  $wd.Visible = $false
  $wd.DisplayAlerts = 0
  Write-Output ("app path: " + $wd.Path)
  $doc = $wd.Documents.Add()
  $doc.Content.Text = "Alpha target line.`rBeta target line.`rTail line.`r"
  # A bookmark over the word "target" in paragraph 1, another over the whole
  # second paragraph, and one over a phrase in the third.
  $text = $doc.Content.Text
  $i = $text.IndexOf("Beta")
  $doc.Bookmarks.Add("alpha", $doc.Range(6, 12))
  $doc.Bookmarks.Add("beta", $doc.Range($i, $i + 4))
  $j = $text.IndexOf("line.", $i)
  $doc.Bookmarks.Add("tail", $doc.Range($j, $j + 5))
  Write-Output ("bookmarks: " + $doc.Bookmarks.Count)
  foreach ($b in $doc.Bookmarks) { Write-Output ("  " + $b.Name + " = " + $b.Start + ".." + $b.End) }
  if ([System.IO.Path]::IsPathRooted($Path)) { $full = $Path } else { $full = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot "..\$Path")) }
  if (Test-Path $full) { Remove-Item $full -Force }
  $doc.SaveAs([ref]$full, [ref]0)
  $doc.Close([ref]0)
  $wd.Quit()
  Write-Output ("saved " + $full)
} catch {
  Write-Output ("ERR " + $_.Exception.Message)
} finally {
  Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
  if ($old) {
    New-Item -Path $clsid -Force | Out-Null
    Set-ItemProperty -Path $clsid -Name '(default)' -Value $old
  }
}
