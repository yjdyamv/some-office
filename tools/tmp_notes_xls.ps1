param([string]$Out = "samples/ms/notes-hlink.xls")

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
    $app.UserName = "ops"
    $wb = $app.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)

    $ws.Range("A1").Value2 = "commented"
    $ws.Range("A1").AddComment("reviewed by ops") | Out-Null

    $ws.Range("B2").Value2 = "link"
    $ws.Hyperlinks.Add($ws.Range("B2"), "https://example.com/page", "", "example page", "click me") | Out-Null

    $ws.Range("A4").Value2 = "second"
    $app.UserName = "qa"
    $ws.Range("A4").AddComment("two lines, second line here") | Out-Null

    $ws.Range("B5").Value2 = "mail"
    $ws.Hyperlinks.Add($ws.Range("B5"), "mailto:someone@example.com", "", "", "mail us") | Out-Null

    $path = [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $Out))
    $wb.SaveAs($path, 1)
    $wb.Close($false)
    $app.Quit()
    Write-Output "app path: $($app.Path)"
    Write-Output "wrote $path"
}
catch {
    Write-Output "FAILED: $($_.Exception.Message)"
}
finally {
    if ($had) { Set-ItemProperty -Path $key -Name '(default)' -Value $old }
    else { Remove-Item -Path $key -Force }
    Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
}
