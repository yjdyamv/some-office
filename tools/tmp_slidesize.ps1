# Make a 16:9 deck and a 4:3 deck so the DocumentAtom of each can be compared.
param()
$ErrorActionPreference = "Stop"
$key = "HKCU:\Software\Classes\CLSID\{0009144D-0000-0000-C000-000000000046}\LocalServer32"
$old = (Get-ItemProperty -Path $key -ErrorAction SilentlyContinue)."(default)"
Remove-Item -Path $key -Force -ErrorAction SilentlyContinue
$root = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot "..\samples\ms"))
$app = New-Object -ComObject PowerPoint.Application
try {
    Write-Output ("app: " + $app.Path)
    foreach ($pair in @(@("widescreen", 15), @("standard", 1))) {
        $name = $pair[0]
        $pres = $app.Presentations.Add($false)
        $pres.PageSetup.SlideSize = $pair[1]
        $slide = $pres.Slides.Add(1, 2)
        $slide.Shapes.Title.TextFrame.TextRange.Text = "Size probe"
        Write-Output ("{0}: w={1} h={2} type={3}" -f $name, $pres.PageSetup.SlideWidth, $pres.PageSetup.SlideHeight, $pres.PageSetup.SlideSize)
        $pres.SaveAs((Join-Path $root ("size-" + $name + ".ppt")), 1)
        $pres.Close()
    }
} finally {
    $app.Quit()
    Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force
    if ($old) { New-Item -Path $key -Force | Out-Null; Set-ItemProperty -Path $key -Name "(default)" -Value $old }
}
