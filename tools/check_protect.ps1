# Read back sheet protection and the locked flag through Excel itself.
param([string]$Path = "samples/written-protected.xls")
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
    Write-Output ("protected: {0}   password: [{1}]" -f $sheet.ProtectContents, $sheet.ProtectionMode)
    Write-Output ("A1 locked: {0}   A2 locked: {1}" -f $sheet.Range("A1").Locked, $sheet.Range("A2").Locked)
    Write-Output ("A2 value: {0}   A2 format: {1}   A2 text: {2}" -f $sheet.Range("A2").Value2, $sheet.Range("A2").NumberFormat, $sheet.Range("A2").Text)
    $book.Close($false)
} finally {
    $app.Quit()
    Start-Sleep -Milliseconds 300
    Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
    if ($old) { New-Item -Path $key -Force | Out-Null; Set-ItemProperty -Path $key -Name "(default)" -Value $old }
}
