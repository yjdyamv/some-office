# Open the files our writers produce with the real applications.
#
# `cargo test` proves the bytes are self-consistent; only Word and Excel can
# prove they are *valid*. Run `cargo run --example mkfiles` first.
$root = Split-Path -Parent $PSScriptRoot
$doc = Join-Path $root 'samples\written.doc'
$xls = Join-Path $root 'samples\written.xls'

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $d = $word.Documents.Open($doc, $false, $true)
    "word   paragraphs=$($d.Paragraphs.Count) chars=$($d.Content.Text.Length)"
    for ($i = 1; $i -le $d.Paragraphs.Count; $i++) {
        "  [$i] $($d.Paragraphs.Item($i).Range.Text.TrimEnd())"
    }
    $d.Close($false)
} catch {
    "word   FAILED: $($_.Exception.Message)"
} finally {
    try { $word.Quit() } catch {}
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $w = $excel.Workbooks.Open($xls, 0, $true)
    $s = $w.Worksheets.Item(1)
    "excel  sheet=$($s.Name) used=$($s.UsedRange.Address(0,0))"
    foreach ($row in 1..4) { "  [$row] $($s.Cells.Item($row,1).Text)" }
    $w.Close($false)
} catch {
    "excel  FAILED: $($_.Exception.Message)"
} finally {
    try { $excel.Quit() } catch {}
}
