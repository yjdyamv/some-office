# Builds samples/ms/validation.xls with Excel 2016: a list validation, a whole
# number range, and a list that points at a cell range.
$clsid = "{00024500-0000-0000-C000-000000000046}"
$key = "HKCU:\Software\Classes\CLSID\$clsid\LocalServer32"
$had = Test-Path $key
$old = if ($had) { (Get-ItemProperty -Path $key)."(default)" } else { $null }
Get-Process et,EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
Remove-Item -Path $key -Force -ErrorAction SilentlyContinue
$out = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot "..\samples\ms\validation.xls"))
try {
  $app = New-Object -ComObject Excel.Application
  "app path: " + $app.Path
  $app.Visible = $false
  $app.DisplayAlerts = $false
  $wb = $app.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1)
  $ws.Name = "Validation"
  $ws.Cells.Item(1, 1).Value2 = "pick a colour"
  $ws.Cells.Item(3, 1).Value2 = 5
  $ws.Cells.Item(5, 1).Value2 = "alpha"
  $ws.Cells.Item(1, 4).Value2 = "alpha"
  $ws.Cells.Item(2, 4).Value2 = "beta"
  $ws.Cells.Item(3, 4).Value2 = "gamma"
  # A1: inline list
  $ws.Range("A1").Validation.Delete()
  $ws.Range("A1").Validation.Add(3, 1, 1, '"red,green,blue"') | Out-Null
  $ws.Range("A1").Validation.InputTitle = "Colour"
  $ws.Range("A1").Validation.InputMessage = "pick one of three"
  $ws.Range("A1").Validation.ErrorTitle = "Not a colour"
  $ws.Range("A1").Validation.ErrorMessage = "red, green or blue"
  # A3: whole number 1..10
  $ws.Range("A3").Validation.Delete()
  $ws.Range("A3").Validation.Add(1, 1, 1, "1", "10") | Out-Null
  # A5: list from a range
  $ws.Range("A5").Validation.Delete()
  $ws.Range("A5").Validation.Add(3, 1, 1, "=D1:D3") | Out-Null
  $ws.Range("A5").Validation.InCellDropdown = $true
  $wb.SaveAs($out, 1)
  $wb.Close($false)
  $app.Quit()
  "wrote " + $out
} catch {
  "ERR " + $_.Exception.Message
} finally {
  if ($had) { New-Item -Path $key -Force | Out-Null
              Set-ItemProperty -Path $key -Name "(default)" -Value $old }
  Get-Process et,EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
}
