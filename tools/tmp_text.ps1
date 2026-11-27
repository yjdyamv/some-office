$ErrorActionPreference = "Stop"
$key = "HKCU:\Software\Classes\CLSID\{91493441-5A91-11CF-8700-00AA0060263B}\LocalServer32"
$old = (Get-ItemProperty -Path $key -ErrorAction SilentlyContinue)."(default)"
try {
    Set-ItemProperty -Path $key -Name "(default)" -Value '"C:\Program Files\Microsoft Office\Root\Office16\POWERPNT.EXE" /automation'
    $ppt = New-Object -ComObject PowerPoint.Application
    foreach ($f in $args) {
        try {
            $pres = $ppt.Presentations.Open((Resolve-Path $f).Path, $true, $false, $false)
            "== $f slides=$($pres.Slides.Count)"
            foreach ($s in $pres.Slides) {
                $t = @()
                foreach ($sh in $s.Shapes) {
                    if ($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1) { $t += $sh.TextFrame.TextRange.Text.Trim() }
                }
                "  slide $($s.SlideIndex): " + ($t -join " / ")
            }
            $pres.Close()
        } catch { "ERR $f : $($_.Exception.Message)" }
    }
    $ppt.Quit()
} finally {
    Set-ItemProperty -Path $key -Name "(default)" -Value $old
    Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force
}
