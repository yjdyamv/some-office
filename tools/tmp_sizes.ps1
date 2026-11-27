$ErrorActionPreference = "Stop"
$key = "HKCU:\Software\Classes\CLSID\{91493441-5A91-11CF-8700-00AA0060263B}\LocalServer32"
$old = (Get-ItemProperty -Path $key)."(default)"
try {
    Set-ItemProperty -Path $key -Name "(default)" -Value '"C:\Program Files\Microsoft Office\Root\Office16\POWERPNT.EXE" /automation'
    $ppt = New-Object -ComObject PowerPoint.Application
    foreach ($k in 1, 3, 4, 6) {
        $pres = $ppt.Presentations.Add(1)          # ppLayoutBlank
        $pres.PageSetup.SlideSize = $k
        $pres.PageSetup.FirstSlideNumber = 7
        $w = $pres.PageSetup.SlideWidth; $h = $pres.PageSetup.SlideHeight
        $pres.SaveAs((Join-Path (Get-Location) "samples\ms\size$k.ppt"), 1)
        "size$k -> ${w}x${h}"
        $pres.Close()
    }
    "OK $($ppt.Path)"
    $ppt.Quit()
} finally {
    Set-ItemProperty -Path $key -Name "(default)" -Value $old
    Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force
}
