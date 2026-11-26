# Open the files our writers produce with the real applications.
#
# `cargo test` proves the bytes are self-consistent; only Word and Excel can
# prove they are *valid*. Run `cargo run --example mkfiles` first.
$root = Split-Path -Parent $PSScriptRoot
$doc = Join-Path $root 'samples\written.doc'
$fmt = Join-Path $root 'samples\written-formatted.doc'
$xls = Join-Path $root 'samples\written.xls'
$fxls = Join-Path $root 'samples\written-formulas.xls'

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
    # One process per file: Word gets into a bad state after a rejection and
    # the next Open would silently return an empty document.
    try { $word.Quit() } catch {}
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    # The same check for the formatted file: only Word can say whether the
    # CHPX/PAPX sprms we encoded mean what we think they mean.
    $d = $word.Documents.Open($fmt, $false, $true)
    "word   formatted: $($d.Paragraphs.Count) paragraphs"
    $r = $d.Paragraphs.Item(1).Range
    "  [1] text=$($r.Text.TrimEnd()) bold=$($r.Font.Bold) size=$($r.Font.Size) colour=$($r.Font.Color) align=$($d.Paragraphs.Item(1).Alignment)"
    $p2 = $d.Paragraphs.Item(2)
    "  [2] left=$($p2.LeftIndent) first=$($p2.FirstLineIndent) italic=$($p2.Range.Characters.Item(7).Font.Italic)"
    "  [3] text=$($d.Paragraphs.Item(3).Range.Text.TrimEnd()) left=$($d.Paragraphs.Item(3).LeftIndent)"
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

# Formulas: Excel must show the formula text *and* compute it, which proves the
# rgce token array is valid and not just a cached number.
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $w = $excel.Workbooks.Open($fxls, 0, $true)
    $excel.CalculateFull()
    $s = $w.Worksheets.Item(1)
    "excel  formulas: sheet=$($s.Name) used=$($s.UsedRange.Address(0,0))"
    foreach ($row in 1..5) {
        "  [$row] $($s.Cells.Item($row,1).Text) | $($s.Cells.Item($row,2).Formula) = $($s.Cells.Item($row,2).Text)"
    }
    $w.Close($false)
} catch {
    "excel  formulas FAILED: $($_.Exception.Message)"
} finally {
    try { $excel.Quit() } catch {}
}

# PowerPoint: write_presentation is self-contained (the template is compiled
# into the crate), so this needs nothing but the generated file.
$ppt = Join-Path $root 'samples\written-scratch.ppt'
if (Test-Path $ppt) {
    $clsid = 'HKCU:\Software\Classes\CLSID\{91493441-5A91-11CF-8700-00AA0060263B}\LocalServer32'
    $saved = (Get-ItemProperty -Path $clsid -Name '(default)' -ErrorAction SilentlyContinue).'(default)'
    try {
        if ($saved) { Set-ItemProperty -Path $clsid -Name '(default)' -Value '"C:\Program Files\Microsoft Office\Root\Office16\POWERPNT.EXE" /automation' }
        $app = New-Object -ComObject PowerPoint.Application
        try {
            $pres = $app.Presentations.Open($ppt, $true, $false, $false)
            "ppt    slides=$($pres.Slides.Count)"
            foreach ($i in 1..$pres.Slides.Count) {
                $texts = @()
                foreach ($j in 1..$pres.Slides.Item($i).Shapes.Count) {
                    $t = $pres.Slides.Item($i).Shapes.Item($j).TextFrame.TextRange.Text
                    if ($t.Trim()) { $texts += $t.Trim() }
                }
                "  s$i $($texts -join ' | ')"
            }
            $pres.Close()
        } catch {
            "ppt    FAILED: $($_.Exception.Message)"
        } finally {
            try { $app.Quit() } catch {}
        }
    } finally {
        if ($saved) { Set-ItemProperty -Path $clsid -Name '(default)' -Value $saved }
    }
}
