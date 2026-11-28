# Read back a .doc's page setup through Word itself.
param([string]$Path = "samples/written-page.doc")
$ErrorActionPreference = "Stop"
$key = "HKCU:\Software\Classes\CLSID\{000209FF-0000-0000-C000-000000000046}\LocalServer32"
$old = (Get-ItemProperty -Path $key -ErrorAction SilentlyContinue)."(default)"
Remove-Item -Path $key -Force -ErrorAction SilentlyContinue
Remove-Item -Path "HKCU:\Software\Microsoft\Office\16.0\Word\Resiliency" -Recurse -Force -ErrorAction SilentlyContinue
$full = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot "..\$Path"))
$word = New-Object -ComObject Word.Application
try {
    Write-Output "app: $($word.Path)"
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Open($full, $false, $true)
    $setup = $doc.PageSetup
    Write-Output ("page: {0} x {1} pt" -f [math]::Round($setup.PageWidth, 1), [math]::Round($setup.PageHeight, 1))
    Write-Output ("margins: L{0} R{1} T{2} B{3} pt" -f [math]::Round($setup.LeftMargin, 1), [math]::Round($setup.RightMargin, 1), [math]::Round($setup.TopMargin, 1), [math]::Round($setup.BottomMargin, 1))
    Write-Output ("orientation: {0} (0 portrait, 1 landscape)" -f $setup.Orientation)
    $doc.Close(0)
} finally {
    $word.Quit()
    Start-Sleep -Milliseconds 300
    Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
    if ($old) {
        New-Item -Path $key -Force | Out-Null
        Set-ItemProperty -Path $key -Name "(default)" -Value $old
    }
}
