# Read the transitions back through PowerPoint itself.
param([string]$Path = "samples/written-transitions.ppt")
$ErrorActionPreference = "Stop"
$key = "HKCU:\Software\Classes\CLSID\{0009144D-0000-0000-C000-000000000046}\LocalServer32"
$old = (Get-ItemProperty -Path $key -ErrorAction SilentlyContinue)."(default)"
Remove-Item -Path $key -Force -ErrorAction SilentlyContinue
$full = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot "..\$Path"))
$app = New-Object -ComObject PowerPoint.Application
try {
    Write-Output ("app: " + $app.Path)
    # PowerPoint refuses Visible = $false; open with WithWindow = $false instead.
    $pres = $app.Presentations.Open($full, $true, $false, $false)
    Write-Output ("slides: {0}" -f $pres.Slides.Count)
    for ($i = 1; $i -le $pres.Slides.Count; $i++) {
        $t = $pres.Slides.Item($i).SlideShowTransition
        Write-Output ("slide {0}: effect={1} advance={2} time={3}" -f $i, $t.EntryEffect, $t.AdvanceOnTime, $t.AdvanceTime)
    }
    $pres.Close()
} finally {
    $app.Quit()
    Start-Sleep -Milliseconds 300
    Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force
    if ($old) { New-Item -Path $key -Force | Out-Null; Set-ItemProperty -Path $key -Name "(default)" -Value $old }
}
