param([string]$Path = "samples/written-filter.xls")
# Open an .xls with the *real* Excel (not WPS) and print its AutoFilter state.
# `$app.Path` is printed on purpose: WPS reports `Application.Name` as
# "Microsoft Excel" / 16.0, so the name alone cannot be trusted.
$clsid = "{00024500-0000-0000-C000-000000000046}"
$key = "HKCU:\Software\Classes\CLSID\$clsid\LocalServer32"
$old = (Get-ItemProperty -Path $key -ErrorAction SilentlyContinue)."(default)"
Get-Process et,EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
Remove-Item -Path $key -Force -ErrorAction SilentlyContinue
try {
  $app = New-Object -ComObject Excel.Application
  "app path: " + $app.Path
  $wb = $app.Workbooks.Open([System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot "..\$Path")))
  $ws = $wb.Worksheets.Item(1)
  "used: " + $ws.UsedRange.Address()
  "autofilter mode: " + $ws.AutoFilterMode
  if ($ws.AutoFilterMode) {
    $af = $ws.AutoFilter
    "range: " + $af.Range.Address()
    "filters: " + $af.Filters.Count
    for ($i = 1; $i -le $af.Filters.Count; $i++) {
      $f = $af.Filters.Item($i)
      "  field {0}: on={1} criteria1={2} operator={3}" -f $i, $f.On, $f.Criteria1, $f.Operator
    }
  }
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
