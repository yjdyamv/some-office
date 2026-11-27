param([string]$Out = "samples\ms\autofilter.xls")
# Write a sheet with an AutoFilter, using the *real* Excel (not WPS).
$clsid = "{00024500-0000-0000-C000-000000000046}"
$key = "HKCU:\Software\Classes\CLSID\$clsid\LocalServer32"
$old = (Get-ItemProperty -Path $key -ErrorAction SilentlyContinue)."(default)"
Get-Process et,EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
Remove-Item -Path $key -Force -ErrorAction SilentlyContinue
try {
  $app = New-Object -ComObject Excel.Application
  $app.Visible = $false
  $app.DisplayAlerts = $false
  "app path: " + $app.Path
  $wb = $app.Workbooks.Add()
  $ws = $wb.Worksheets.Item(1)
  $ws.Name = "Filtered"
  $data = @(
    @("name", "kind", "size"),
    @("alpha", "red", 10),
    @("beta", "green", 20),
    @("gamma", "red", 30),
    @("delta", "blue", 40)
  )
  for ($r = 0; $r -lt $data.Count; $r++) {
    for ($c = 0; $c -lt 3; $c++) { $ws.Cells.Item($r + 1, $c + 1).Value2 = $data[$r][$c] }
  }
  $range = $ws.Range("A1:C5")
  $range.AutoFilter(1) | Out-Null
  # Filter the "kind" column down to one value.
  $ws.Range("B1").AutoFilter(2, "red") | Out-Null
  $path = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot "..\$Out"))
  $wb.SaveAs($path, 1)
  "saved " + $path
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
