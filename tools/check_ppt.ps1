param([string[]]$Files)
$clsid = "{91493441-5A91-11CF-8700-00AA0060263B}"
$key = "HKCU:\Software\Classes\CLSID\$clsid\LocalServer32"
$old = (Get-ItemProperty -Path $key -ErrorAction SilentlyContinue)."(default)"
Set-ItemProperty -Path $key -Name "(default)" -Value '"C:\Program Files\Microsoft Office\Root\Office16\POWERPNT.EXE" /automation'
try {
  foreach ($f in $Files) {
    $app = New-Object -ComObject PowerPoint.Application
    try {
      $full = (Resolve-Path $f).Path
      $p = $app.Presentations.Open($full, $true, $false, $false)
      "OK  $([System.IO.Path]::GetFileName($f)) slides=$($p.Slides.Count)"
      foreach ($s in $p.Slides) {
        $t = @()
        foreach ($sh in $s.Shapes) { $t += ("type=" + $sh.Type + " " + [math]::Round($sh.Width,1) + "x" + [math]::Round($sh.Height,1)) }
        "  slide $($s.SlideIndex): " + ($t -join " | ")
      }
      $p.Close()
    } catch { "ERR $([System.IO.Path]::GetFileName($f)) $($_.Exception.Message)" }
    finally { $app.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null }
    Start-Sleep -Milliseconds 300
  }
} finally {
  Set-ItemProperty -Path $key -Name "(default)" -Value $old
  Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force
}
