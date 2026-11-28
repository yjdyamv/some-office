$ErrorActionPreference = "Stop"
$key = "HKCU:\Software\Classes\CLSID\{0009144D-0000-0000-C000-000000000046}\LocalServer32"
$old = (Get-ItemProperty -Path $key -ErrorAction SilentlyContinue)."(default)"
Remove-Item -Path $key -Force -ErrorAction SilentlyContinue
$root = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot "..\samples\ms"))
$effects = @{ cut = 257; fade = 1793; wipedown = 1281; dissolve = 1537; blinds = 769; checker = 1795; boxout = 1284; random = 513; splitout = 1025 }
$app = New-Object -ComObject PowerPoint.Application
try {
    Write-Output ("app: " + $app.Path)
    foreach ($name in $effects.Keys) {
        $pres = $app.Presentations.Add(0)
        $slide = $pres.Slides.Add(1, 2)
        $slide.Shapes.Title.TextFrame.TextRange.Text = "T " + $name
        $slide.SlideShowTransition.EntryEffect = $effects[$name]
        $pres.SaveAs((Join-Path $root "tr-$name.ppt"), 1)
        $pres.Close()
    }
} finally {
    $app.Quit()
    Start-Sleep -Milliseconds 300
    Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force
    if ($old) { New-Item -Path $key -Force | Out-Null; Set-ItemProperty -Path $key -Name "(default)" -Value $old }
}
