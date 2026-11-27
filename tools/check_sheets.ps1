param([string]$File)
# Open an .xls with the *real* Excel (not WPS) and list every worksheet with its
# A1 value, so a multi-sheet workbook can be checked end to end.
# `$app.Path` is printed on purpose: WPS reports `Application.Name` as
# "Microsoft Excel" / version 16.0, so the name alone cannot be trusted.
$clsid = "{00024500-0000-0000-C000-000000000046}"
$key = "HKCU:\Software\Classes\CLSID\$clsid\LocalServer32"
$old = (Get-ItemProperty -Path $key)."(default)"
Get-Process et,EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
Set-ItemProperty -Path $key -Name "(default)" -Value '"C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE" /automation'
try {
  $app = New-Object -ComObject Excel.Application
  "app path: " + $app.Path
  $wb = $app.Workbooks.Open([System.IO.Path]::GetFullPath($File))
  "sheets: " + $wb.Worksheets.Count
  foreach ($ws in $wb.Worksheets) {
    $a1 = $ws.Cells.Item(1, 1)
    "  {0}: A1={1} text={2} fmt={3} width={4}" -f $ws.Name, $a1.Value2, $a1.Text, $a1.NumberFormat, $ws.Columns.Item(1).ColumnWidth
  }
  $wb.Close($false)
  $app.Quit()
} catch {
  "ERR " + $_.Exception.Message
} finally {
  Set-ItemProperty -Path $key -Name "(default)" -Value $old
  Get-Process et,EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
}
