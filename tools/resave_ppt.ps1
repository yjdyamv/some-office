param([string]$Path = "samples/written-ppt-metadata.ppt")
# Open a deck in real PowerPoint and save it again: the property sets PowerPoint
# keeps are the ones it understood.
$ErrorActionPreference = 'Continue'
$clsid = 'HKCU:\Software\Classes\CLSID\{00020906-0000-0000-C000-000000000046}\LocalServer32'
$old = $null
if (Test-Path $clsid) { $old = (Get-ItemProperty $clsid).'(default)' }
if ($old) { Remove-Item -Path $clsid -Force }
Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force
try {
  $pp = New-Object -ComObject PowerPoint.Application
  Write-Output ("app path: " + $pp.Path)
  if ([System.IO.Path]::IsPathRooted($Path)) { $full = $Path } else { $full = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot "..\$Path")) }
  $target = [System.IO.Path]::ChangeExtension($full, ".resaved.ppt")
  if (Test-Path $target) { Remove-Item $target -Force }
  $deck = $pp.Presentations.Open($full, $true, $false, $false)
  Write-Output ("opened slides=" + $deck.Slides.Count)
  # 1 = ppSaveAsPresentation (.ppt, keeps the property sets).
  $deck.SaveAs($target, 1)
  $deck.Close()
  $pp.Quit()
  Write-Output ("resaved to " + $target)
} catch {
  Write-Output ("ERR " + $_.Exception.Message)
} finally {
  Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force
  if ($old) {
    New-Item -Path $clsid -Force | Out-Null
    Set-ItemProperty -Path $clsid -Name '(default)' -Value $old
  }
}
