$ErrorActionPreference = "Stop"
$key = "HKCU:\Software\Classes\CLSID\{91493441-5A91-11CF-8700-00AA0060263B}\LocalServer32"
$old = (Get-ItemProperty -Path $key -ErrorAction SilentlyContinue)."(default)"
try {
    Set-ItemProperty -Path $key -Name "(default)" -Value '"C:\Program Files\Microsoft Office\Root\Office16\POWERPNT.EXE" /automation'
    $ppt = New-Object -ComObject PowerPoint.Application
    $pres = $ppt.Presentations.Add(1)
    $pres.PageSetup.FirstSlideNumber = 5
    $pres.PageSetup.SlideSize = 2   # ppSlideSizeOnScreen16x9
    $pres.SaveAs((Join-Path (Get-Location) "samples\ms\fsn.ppt"), 1)
    "OK $($ppt.Path)"
    $pres.Close(); $ppt.Quit()
} finally {
    Set-ItemProperty -Path $key -Name "(default)" -Value $old
    Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force
}
