param([string]$Path = "samples/written-links.xls")
# Open an .xls with the *real* Excel (not WPS) and print its hyperlinks.
# `$app.Path` is printed on purpose: WPS reports `Application.Name` as
# "Microsoft Excel" / 16.0, so the name alone cannot be trusted.
$clsid = "{00024500-0000-0000-C000-000000000046}"
$key = "HKCU:\Software\Classes\CLSID\$clsid\LocalServer32"
$old = (Get-ItemProperty -Path $key)."(default)"
Get-Process et,EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
# Deleting the per-user override is what actually makes COM pick the machine
# registration (Microsoft Excel); rewriting it in place is ignored here.
Remove-Item -Path $key -Force
try {
  $app = New-Object -ComObject Excel.Application
  "app path: " + $app.Path
  $wb = $app.Workbooks.Open([System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot "..\$Path")))
  $ws = $wb.Worksheets.Item(1)
  $used = $ws.UsedRange
  "used: " + $used.Address()
  foreach ($cell in $used.Cells) {
    $text = $cell.Text
    if ($null -eq $text -or "$text" -eq "") { continue }
    if ($cell.Hyperlinks.Count -gt 0) {
      $link = $cell.Hyperlinks.Item(1)
      "{0}: text={1} address={2} type={3}" -f $cell.Address(), $text, $link.Address, $link.Type
    } else {
      "{0}: text={1} (no link)" -f $cell.Address(), $text
    }
  }
  $wb.Close($false)
  $app.Quit()
} catch {
  "ERR " + $_.Exception.Message
} finally {
  New-Item -Path $key -Force | Out-Null
  New-Item -Path $key -Force | Out-Null
  Set-ItemProperty -Path $key -Name "(default)" -Value $old
  Get-Process et,EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
}
