# A protected sheet (password "abc") with one unlocked cell.
$ErrorActionPreference = "Stop"
$key = "HKCU:\Software\Classes\CLSID\{00024500-0000-0000-C000-000000000046}\LocalServer32"
$old = (Get-ItemProperty -Path $key -ErrorAction SilentlyContinue)."(default)"
Remove-Item -Path $key -Force -ErrorAction SilentlyContinue
$root = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot "..\samples\ms"))
$app = New-Object -ComObject Excel.Application
try {
    Write-Output "app: $($app.Path)"
    $app.Visible = $false
    $app.DisplayAlerts = $false
    $book = $app.Workbooks.Add()
    $sheet = $book.Worksheets.Item(1)
    $sheet.Range("A1:B2").Value2 = @(@(1,2), @(3,4))
    $sheet.Range("B2").Locked = $false
    $sheet.Protect("abc", $true, $true, $true)
    $path = Join-Path $root "protected.xls"
    $book.SaveAs([string]$path, 1)
    Write-Output ("protected: {0}  B2 locked: {1}" -f $sheet.ProtectContents, $sheet.Range("B2").Locked)
    $book.Close($false)
} finally {
    $app.Quit()
    Start-Sleep -Milliseconds 300
    Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
    if ($old) { New-Item -Path $key -Force | Out-Null; Set-ItemProperty -Path $key -Name "(default)" -Value $old }
}
