param([string]$In = "samples/written-fills.xls")

$key = "HKCU:\Software\Classes\CLSID\{00024500-0000-0000-C000-000000000046}\LocalServer32"
$old = (Get-ItemProperty -Path $key -ErrorAction SilentlyContinue).'(default)'
New-Item -Force -Path $key | Out-Null
# A per-user override wins over the machine registration, so delete it
Remove-Item -Path $key -Force -ErrorAction SilentlyContinue

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
    # Put back whatever the user had. The key itself has to be recreated first:
    # Set-ItemProperty cannot write into a path that does not exist.
    if ($old) {
        New-Item -Path $key -Force | Out-Null
        Set-ItemProperty -Path $key -Name '(default)' -Value $old
    }
    else { Remove-Item -Path $key -Force -ErrorAction SilentlyContinue }
    Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
}
