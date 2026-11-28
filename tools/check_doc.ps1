# Check a .doc written by office-doc: does Word open it, and does it see the picture?
$ErrorActionPreference = 'Continue'
$path = (Resolve-Path $args[0]).Path
$clsid = 'HKCU:\Software\Classes\CLSID\{000209FF-0000-0000-C000-000000000046}\LocalServer32'
$exe = '"C:\Program Files\Microsoft Office\Root\Office16\WINWORD.EXE" /Automation'
$backup = $null
if (Test-Path $clsid) { $backup = (Get-ItemProperty $clsid).'(default)' }
New-Item -Path $clsid -Force | Out-Null
Set-ItemProperty -Path $clsid -Name '(default)' -Value $exe
try {
    # Word remembers every path it failed to open in `Resiliency\DisabledItems` and
    # then hangs on that path forever, even after the file is fixed. Clear the
    # list before each run so a repaired sample can be checked again.
    Remove-Item -Path 'HKCU:\Software\Microsoft\Office\16.0\Word\Resiliency' -Recurse -Force -ErrorAction SilentlyContinue
    $wd = New-Object -ComObject Word.Application
    $wd.Visible = $false
    $wd.DisplayAlerts = 0
    $doc = $wd.Documents.Open($path, $false, $true)
    Write-Host ("paras={0} words={1} chars={2} shapes={3} inline={4}" -f `
        $doc.Paragraphs.Count, $doc.Words.Count, $doc.Characters.Count, `
        $doc.Shapes.Count, $doc.InlineShapes.Count)
    foreach ($s in $doc.InlineShapes) {
        Write-Host ("  inline shape type={0} w={1:N1} h={2:N1}" -f $s.Type, $s.Width, $s.Height)
    }
    foreach ($p in $doc.Paragraphs) {
        Write-Host ("  [{0}] {1}" -f $p.Range.ListFormat.ListString, $p.Range.Text.Trim())
    }
    $doc.Close($false)
    $wd.Quit()
} catch {
    Write-Host "ERR $($_.Exception.Message)"
} finally {
    if ($backup) { Set-ItemProperty -Path $clsid -Name '(default)' -Value $backup }
    Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
}
