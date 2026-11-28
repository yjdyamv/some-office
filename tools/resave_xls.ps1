param([string]$Path = "samples/written-xls-metadata.xls")
# Open a workbook in real Excel and save it again: the property sets Excel
# keeps are the ones it understood.
$ErrorActionPreference = 'Continue'
$clsid = 'HKCU:\Software\Classes\CLSID\{00024500-0000-0000-C000-000000000046}\LocalServer32'
$old = $null
if (Test-Path $clsid) { $old = (Get-ItemProperty $clsid).'(default)' }
if ($old) { Remove-Item -Path $clsid -Force }
Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
try {
  $xl = New-Object -ComObject Excel.Application
  $xl.Visible = $false
  $xl.DisplayAlerts = $false
  Write-Output ("app path: " + $xl.Path)
  if ([System.IO.Path]::IsPathRooted($Path)) { $full = $Path } else { $full = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot "..\$Path")) }
  $target = [System.IO.Path]::ChangeExtension($full, ".resaved.xls")
  if (Test-Path $target) { Remove-Item $target -Force }
  $book = $xl.Workbooks.Open($full)
  Write-Output ("opened sheets=" + $book.Sheets.Count + " cells=" + $book.Sheets.Item(1).Range("A1").Text)
  # 56 = xlExcel8, the BIFF8 format that keeps the property sets.
  $book.SaveAs($target, 56)
  $book.Close($false)
  $xl.Quit()
  Write-Output ("resaved to " + $target)
} catch {
  Write-Output ("ERR " + $_.Exception.Message)
} finally {
  Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
  if ($old) {
    New-Item -Path $clsid -Force | Out-Null
    Set-ItemProperty -Path $clsid -Name '(default)' -Value $old
  }
}
