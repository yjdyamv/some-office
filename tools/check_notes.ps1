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
      "OK  $([System.IO.Path]::GetFileName($f)) slides=$($p.Slides.Count) path=$($app.Path)"
      foreach ($s in $p.Slides) {
        $n = ""
        try {
          foreach ($sh in $s.NotesPage.Shapes) {
            if ($sh.HasTextFrame -eq -1 -and $sh.TextFrame.HasText -eq -1) { $n += "[" + $sh.TextFrame.TextRange.Text + "]" }
          }
        } catch { $n = "<err $($_.Exception.Message)>" }
        "  slide $($s.SlideIndex) notes: $n"
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
