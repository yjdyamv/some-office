param([string]$Out = "samples/ms/layout.xls")

$clsid = "{00024500-0000-0000-C000-000000000046}"
$key = "HKCU:\Software\Classes\CLSID\$clsid\LocalServer32"
$excel = "C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE"
$existed = Test-Path $key
$old = $null
if ($existed) { $old = (Get-ItemProperty -Path $key)."(default)" }
"key existed before: $existed"
Get-Process et, EXCEL, wps -ErrorAction SilentlyContinue | Stop-Process -Force
New-Item -Path $key -Force | Out-Null
Set-ItemProperty -Path $key -Name "(default)" -Value ('"' + $excel + '" /automation')
"patched: " + (Get-ItemProperty -Path $key)."(default)"
try {
    $app = New-Object -ComObject Excel.Application
    "app path: " + $app.Path
    $app.Visible = $false
    $app.DisplayAlerts = $false
    $wb = $app.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = "Layout"
    $ws.Range("A1").Value2 = "merged title"
    $ws.Range("A1:C1").Merge()
    $ws.Range("A3").Value2 = "left"
    $ws.Range("B3").Value2 = 42
    $ws.Range("C3").Value2 = "wide"
    $ws.Columns.Item(1).ColumnWidth = 20
    $ws.Columns.Item(2).ColumnWidth = 5
    $ws.Columns.Item(3).ColumnWidth = 30
    $ws.Rows.Item(1).RowHeight = 40
    $ws.Rows.Item(3).RowHeight = 22.5
    $ws.Activate()
    $ws.Range("B2").Select()
    $app.ActiveWindow.FreezePanes = $true

    $full = [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $Out))
    if (Test-Path $full) { Remove-Item $full -Force }
    $wb.SaveAs($full, 56)
    "OK  $full"
    $wb.Close($false)
    $app.Quit()
} catch {
    "ERR " + $_.Exception.Message
} finally {
    if ($existed) {
        New-Item -Path $key -Force | Out-Null
        Set-ItemProperty -Path $key -Name "(default)" -Value $old
    } else {
        Remove-Item -Path $key -Force -ErrorAction SilentlyContinue
    }
    Get-Process et, EXCEL, wps -ErrorAction SilentlyContinue | Stop-Process -Force
}
