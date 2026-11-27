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
  $wb = $app.Workbooks.Open([System.IO.Path]::GetFullPath($File))
  $ws = $wb.Worksheets.Item(1)
  "sheet: " + $ws.Name
  for ($r = 1; $r -le $Rows; $r++) {
    foreach ($col in 1, 2) {
      $c = $ws.Cells.Item($r, $col)
      if ($null -eq $c.Value2) { continue }
      $name = [char](64 + $col)
      "  {0}{1} value={2} text={3} fmt={4} formula={5}" -f $name, $r, $c.Value2, $c.Text, $c.NumberFormat, $c.Formula
    }
  }
  $wb.Close($false)
  $app.Quit()
} catch {
  "ERR " + $_.Exception.Message
} finally {
  Set-ItemProperty -Path $key -Name "(default)" -Value $old
  Get-Process et,EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
}
