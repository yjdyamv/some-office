param([string]$File = "samples/written-borders.xls")

# Excel, not WPS: WPS claims the same ProgID and even the same Name, so point the
# CLSID at the real EXCEL.EXE for the duration of the run and restore it after.
$key = "HKCU:\Software\Classes\CLSID\{00024500-0000-0000-C000-000000000046}\LocalServer32"
$had = Test-Path $key
$old = if ($had) { (Get-ItemProperty $key).'(default)' } else { $null }
New-Item -Force -Path $key | Out-Null
# A per-user override wins over the machine registration, so delete it
Remove-Item -Path $key -Force -ErrorAction SilentlyContinue

try {
    $app = New-Object -ComObject Excel.Application
    $app.Visible = $false
    $app.DisplayAlerts = $false
    $path = [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $File))
    $wb = $app.Workbooks.Open($path)
    $ws = $wb.Worksheets.Item(1)
    Write-Output "app path: $($app.Path)"

    foreach ($addr in @("A1", "A3", "A5", "A7", "A9", "A11", "C1")) {
        $r = $ws.Range($addr)
        $edges = @()
        foreach ($e in @(7, 10, 8, 9, 5, 6)) {
            $b = $r.Borders.Item($e)
            if ($b.LineStyle -ne -4142) {
                $edges += ("{0}=style{1}/weight{2}/colour{3}" -f $e, $b.LineStyle, $b.Weight, $b.Color)
            }
        }
        $align = "h=$($r.HorizontalAlignment) v=$($r.VerticalAlignment) wrap=$($r.WrapText) indent=$($r.IndentLevel)"
        Write-Output ("  {0} {1,-22} {2} [{3}]" -f $addr, $r.Text, $align, ($edges -join " "))
    }

    $wb.Close($false)
    $app.Quit()
}
finally {
    if ($had) { Set-ItemProperty -Path $key -Name '(default)' -Value $old }
    else { Remove-Item -Path $key -Force }
    Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
}
