param([string]$Out = "samples/ms/palette-custom.xls")
# Change the workbook palette itself (Workbook.Colors) so that Excel has to
# write a PALETTE record, then read its bytes to settle the RGB/BGR order.
$clsid = "{00024500-0000-0000-C000-000000000046}"
$key = "HKCU:\Software\Classes\CLSID\$clsid\LocalServer32"
$old = (Get-ItemProperty -Path $key -ErrorAction SilentlyContinue)."(default)"
Get-Process et,EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
Remove-Item -Path $key -Force -ErrorAction SilentlyContinue
try {
  $app = New-Object -ComObject Excel.Application
  $app.DisplayAlerts = $false
  "app path: " + $app.Path
  $wb = $app.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1)
  # BGR packed ints: 0x030201 -> R=1 G=2 B=3
  $wb.Colors(8) = 0x00030201
  $wb.Colors(9) = 0x00060504
  $ws.Cells.Item(1, 1).Value2 = "palette"
  $path = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot "..\$Out"))
  $wb.SaveAs($path, 1)
  "colours 8/9: " + $wb.Colors(8) + " " + $wb.Colors(9)
  $wb.Close($false)
  $app.Quit()
} catch {
  "ERR " + $_.Exception.Message
} finally {
  if ($old) {
    New-Item -Path $key -Force | Out-Null
    Set-ItemProperty -Path $key -Name "(default)" -Value $old
  }
  Get-Process et,EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
}