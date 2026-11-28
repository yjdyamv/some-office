# Ask Word to write a document with a bookmark and an internal hyperlink.
$ErrorActionPreference = 'Continue'
$clsid = 'HKCU:\Software\Classes\CLSID\{000209FF-0000-0000-C000-000000000046}\LocalServer32'
$old = $null
if (Test-Path $clsid) { $old = (Get-ItemProperty $clsid).'(default)' }
if ($old) { Remove-Item -Path $clsid -Force }
Remove-Item -Path 'HKCU:\Software\Microsoft\Office\16.0\Word\Resiliency' -Recurse -Force -ErrorAction SilentlyContinue
Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
try {
  $wd = New-Object -ComObject Word.Application
  $wd.Visible = $false
  Write-Output ("app path: " + $wd.Path)
  $doc = $wd.Documents.Add()
  $doc.Content.Text = "Jump here. See anchor now.`r"
  $target = $doc.Range(15, 21)          # the word "anchor"
  $doc.Bookmarks.Add("alpha", $target) | Out-Null
  $anchor = $doc.Range(0, 4)            # the word "Jump"
  $h = $doc.Hyperlinks.Add($anchor, "", "alpha", "", "jump")
  Write-Output ("added link subaddress=[" + $h.SubAddress + "]")
  $full = [System.IO.Path]::GetFullPath("C:\Users\admin\Desktop\some-office\samples\ms\hyperlink-internal.doc")
  $doc.SaveAs([ref]$full, [ref]0)
  Write-Output ("saved " + $full)
  $doc.Close($false); $wd.Quit()
} catch { Write-Output ("ERR " + $_.Exception.Message) }
finally {
  Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
  if ($old) { New-Item -Path $clsid -Force | Out-Null; Set-ItemProperty -Path $clsid -Name '(default)' -Value $old }
}
