param([string]$File)
# Open an .xls with the *real* Excel (not WPS) and print each row's font.
# `$app.Path` is printed on purpose: WPS reports `Application.Name` as
# "Microsoft Excel" / version 16.0, so the name alone cannot be trusted.
$clsid = "{00024500-0000-0000-C000-000000000046}"
$key = "HKCU:\Software\Classes\CLSID\$clsid\LocalServer32"
$old = (Get-ItemProperty -Path $key)."(default)"
Get-Process et,EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
Set-ItemProperty -Path $key -Name "(default)" -Value '"C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE" /automation'
try {
  $app = New-Object -ComObject Excel.Application
  "app path: " + $app.Path
  $wb = $app.Workbooks.Open([System.IO.Path]::GetFullPath($File))
  $ws = $wb.Worksheets.Item(1)
  for ($r = 1; $r -le 6; $r++) {
    $c = $ws.Cells.Item($r, 1)
    if ($null -eq $c.Value2) { continue }
    $f = $c.Font
    $colour = if ($f.Color -lt 0) { "automatic" } else { "#{0:x6}" -f ($f.Color -band 0xFFFFFF) }
    "  A{0} text={1} bold={2} italic={3} size={4} colour={5} name={6} fmt={7}" -f `
      $r, $c.Text, $f.Bold, $f.Italic, $f.Size, $colour, $f.Name, $c.NumberFormat
  }
  $wb.Close($false)
  $app.Quit()
} catch {
  "ERR " + $_.Exception.Message
} finally {
  Set-ItemProperty -Path $key -Name "(default)" -Value $old
  Get-Process et,EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
}
