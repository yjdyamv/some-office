# Read back hidden rows and columns through Excel itself.
param([string]$Path = "samples/written-hidden.xls")
$ErrorActionPreference = "Stop"
$key = "HKCU:\Software\Classes\CLSID\{00024500-0000-0000-C000-000000000046}\LocalServer32"
$old = (Get-ItemProperty -Path $key -ErrorAction SilentlyContinue)."(default)"
Remove-Item -Path $key -Force -ErrorAction SilentlyContinue
$full = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot "..\$Path"))
$app = New-Object -ComObject Excel.Application
try {
    Write-Output "app: $($app.Path)"
    $app.Visible = $false
    $app.DisplayAlerts = $false
    $book = $app.Workbooks.Open($full, 0, $true)
    $sheet = $book.Worksheets.Item(1)
    Write-Output ("sheet: {0}" -f $sheet.Name)
    Write-Output ("column A hidden: {0}   column B hidden: {1}   width B: {2}" -f $sheet.Columns.Item(1).Hidden, $sheet.Columns.Item(2).Hidden, $sheet.Columns.Item(2).ColumnWidth)
    Write-Output ("row 1 hidden: {0}   row 2 hidden: {1}" -f $sheet.Rows.Item(1).Hidden, $sheet.Rows.Item(2).Hidden)
    $book.Close($false)
} finally {
    $app.Quit()
    Start-Sleep -Milliseconds 300
    Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
    if ($old) { New-Item -Path $key -Force | Out-Null; Set-ItemProperty -Path $key -Name "(default)" -Value $old }
}
