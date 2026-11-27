param([string]$In = "samples/written-fills.xls")

$key = "HKCU:\Software\Classes\CLSID\{00024500-0000-0000-C000-000000000046}\LocalServer32"
$had = Test-Path $key
$old = if ($had) { (Get-ItemProperty $key).'(default)' } else { $null }
New-Item -Force -Path $key | Out-Null
Set-ItemProperty -Path $key -Name '(default)' -Value '"C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE" /automation'

try {
    $app = New-Object -ComObject Excel.Application
    $app.Visible = $false
    $app.DisplayAlerts = $false
    $wb = $app.Workbooks.Open([System.IO.Path]::GetFullPath((Join-Path (Get-Location) $In)))
    $ws = $wb.Worksheets.Item(1)
    Write-Output "app path: $($app.Path)"

    foreach ($row in 1, 3, 5, 7, 9) {
        $cell = $ws.Cells.Item($row, 1)
        $i = $cell.Interior
        $fore = if ($i.Color -lt 0) { $i.Color + 16777216 } else { $i.Color }
        Write-Output ("  A{0} {1,-14} pattern={2,-3} fore={3,-8} patternColor={4}" -f `
            $row, $cell.Text, $i.Pattern, $fore, $i.PatternColor)
    }

    $wb.Close($false)
    $app.Quit()
}
finally {
    if ($had) { Set-ItemProperty -Path $key -Name '(default)' -Value $old }
    else { Remove-Item -Path $key -Force }
    Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
}
