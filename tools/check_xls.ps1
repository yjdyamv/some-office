param([string]$File, [int]$Rows = 4)
# Open an .xls with the *real* Excel (not WPS) and print the first column.
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
  $wb = $app.Workbooks.Open($File)
  $ws = $wb.Worksheets.Item(1)
  "sheet: " + $ws.Name
  for ($r = 1; $r -le $Rows; $r++) {
    $c = $ws.Cells.Item($r, 1)
    if ($null -eq $c.Value2) { continue }
    "  A{0} value={1} text={2} fmt={3}" -f $r, $c.Value2, $c.Text, $c.NumberFormat
  }
  $wb.Close($false)
  $app.Quit()
} catch {
  "ERR " + $_.Exception.Message
} finally {
  Set-ItemProperty -Path $key -Name "(default)" -Value $old
  Get-Process et,EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
}
