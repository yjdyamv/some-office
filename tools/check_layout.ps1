param([string]$File)
# Read back the sheet furniture the writer emits: merged ranges, column widths,
# row heights and frozen panes. Uses the *real* Excel — WPS hijacks the
# `Excel.Application` CLSID, so point it at EXCEL.EXE for the duration (the same
# trick as tools/check_xls.ps1) and print `$app.Path` to prove which one ran.
$clsid = "{00024500-0000-0000-C000-000000000046}"
$key = "HKCU:\Software\Classes\CLSID\$clsid\LocalServer32"
$existed = Test-Path $key
$old = $null
if ($existed) { $old = (Get-ItemProperty -Path $key)."(default)" }
Get-Process et, EXCEL, wps -ErrorAction SilentlyContinue | Stop-Process -Force
New-Item -Path $key -Force | Out-Null
Set-ItemProperty -Path $key -Name "(default)" -Value '"C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE" /automation'
try {
    $app = New-Object -ComObject Excel.Application
    $app.Visible = $false
    $app.DisplayAlerts = $false
    $wb = $app.Workbooks.Open((Resolve-Path $File))
    $ws = $wb.Worksheets.Item(1)
    "OK  {0} path={1}" -f (Split-Path $File -Leaf), $app.Path
    "  sheet: {0}" -f $ws.Name
    $merged = @()
    foreach ($area in $ws.UsedRange.MergeArea) { }
    $cell = $ws.Range("A1")
    "  A1 text: {0}" -f $cell.Text
    "  A1 merge: {0}" -f $cell.MergeArea.Address($false, $false)
    "  merged count: {0}" -f $ws.UsedRange.MergeCells
    foreach ($col in "A", "B", "C") {
        "  width {0}: {1}" -f $col, [math]::Round($ws.Columns.Item($col).ColumnWidth, 2)
    }
    foreach ($row in 1, 3) {
        "  height {0}: {1}" -f $row, [math]::Round($ws.Rows.Item($row).RowHeight, 2)
    }
    "  freeze panes: {0} splitRow={1} splitCol={2}" -f `
        $app.ActiveWindow.FreezePanes, $app.ActiveWindow.SplitRow, $app.ActiveWindow.SplitColumn
    $wb.Close($false)
    $app.Quit()
} catch {
    "ERR " + $_.Exception.Message
} finally {
    if ($existed) {
        Set-ItemProperty -Path $key -Name "(default)" -Value $old
    } else {
        Remove-Item -Path $key -Force -ErrorAction SilentlyContinue
    }
    Get-Process et, EXCEL, wps -ErrorAction SilentlyContinue | Stop-Process -Force
}
