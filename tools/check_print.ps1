# Read back a workbook's print setup through Excel itself.
param([string]$Path = "samples/written-print.xls")
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
    $setup = $sheet.PageSetup
    Write-Output ("sheet: {0}" -f $sheet.Name)
    Write-Output ("orientation: {0} (1 portrait, 2 landscape)" -f $setup.Orientation)
    Write-Output ("paper: {0} (9 = A4)" -f $setup.PaperSize)
    Write-Output ("margins pt: L{0} R{1} T{2} B{3}" -f [math]::Round($setup.LeftMargin, 2), [math]::Round($setup.RightMargin, 2), [math]::Round($setup.TopMargin, 2), [math]::Round($setup.BottomMargin, 2))
    Write-Output ("header: [{0}]" -f $setup.CenterHeader)
    Write-Output ("footer: [{0}]" -f $setup.CenterFooter)
    Write-Output ("print area: [{0}]" -f $setup.PrintArea)
    $book.Close($false)
} finally {
    $app.Quit()
    Start-Sleep -Milliseconds 300
    Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
    if ($old) { New-Item -Path $key -Force | Out-Null; Set-ItemProperty -Path $key -Name "(default)" -Value $old }
}
