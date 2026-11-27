param([string]$Path = "samples/written-validation.xls")
# Open an .xls with the *real* Excel (not WPS) and print its data validation
# rules. `$app.Path` is printed on purpose: WPS reports `Application.Name` as
# "Microsoft Excel" / 16.0, so the name alone cannot be trusted.
$clsid = "{00024500-0000-0000-C000-000000000046}"
$key = "HKCU:\Software\Classes\CLSID\$clsid\LocalServer32"
$old = (Get-ItemProperty -Path $key -ErrorAction SilentlyContinue)."(default)"
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
    if ($cell.Validation.Type -ne -4142) {   # xlValidateInputOnly
      $v = $cell.Validation
      "{0}: type={1} op={2} f1={3} f2={4} dropdown={5} blank={6}" -f `
        $cell.Address(), $v.Type, $v.Operator, $v.Formula1, $v.Formula2, `
        $v.InCellDropdown, $v.IgnoreBlank
      "    input: [{0}] {1}" -f $v.InputTitle, $v.InputMessage
      "    error: [{0}] {1} style={2}" -f $v.ErrorTitle, $v.ErrorMessage, $v.ErrorStyle
    } else {
      "{0}: text={1} (no validation)" -f $cell.Address(), $text
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
