$clsid = "{91493441-5A91-11CF-8700-00AA0060263B}"
$key = "HKCU:\Software\Classes\CLSID\$clsid\LocalServer32"
$old = (Get-ItemProperty -Path $key -ErrorAction SilentlyContinue)."(default)"
Set-ItemProperty -Path $key -Name "(default)" -Value '"C:\Program Files\Microsoft Office\Root\Office16\POWERPNT.EXE" /automation'
try {
  $app = New-Object -ComObject PowerPoint.Application
  $src = (Resolve-Path "samples/void.pptx").Path
  $dst = (Join-Path (Get-Location) "samples/void-ms.ppt")
  if (Test-Path $dst) { Remove-Item $dst -Force }
  $p = $app.Presentations.Open($src, $false, $false, $false)
  "opened: slides=$($p.Slides.Count) layouts=$($p.SlideMaster.CustomLayouts.Count)"
  $p.SaveAs($dst, 1)
  $p.Close()
  $app.Quit()
  "saved $dst"
} catch { "ERR $($_.Exception.Message)" }
finally {
  Set-ItemProperty -Path $key -Name "(default)" -Value $old
  Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force
}
