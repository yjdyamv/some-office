<#
.SYNOPSIS
    用**微软自己**的 Office 生成参考样本（samples/ms/）。

.DESCRIPTION
    这台机器上 HKCU 的 CLSID 注册指向 WPS，HKCU 又压过 HKLM，所以
    `New-Object -ComObject Word.Application` 永远落到 wps.exe 上。WPS 写出来的
    文件是合法的，但它会忽略一些参数（例如 `xlExcel9795` 被无视，照样写 BIFF8），
    于是「BIFF5 / XOR / CryptoAPI」这些分支就拿不到真值。

    本脚本临时把三个 CLSID 的 HKCU `LocalServer32` 指向 Office16，生成完立刻还原
    （备份存在内存里，`finally` 里恢复；WPS 自己的注册项不动）。

    产出：
        hello.doc        ASCII + 中文 + 全角符号
        formatted.doc    粗/斜/下划线/字号/颜色 + 表格 + 脚注 + 页眉
        locked.doc       密码打开（微软写的是 RC4 CryptoAPI 128 位）
        formulas.xls     16 条公式 + 4 种数字格式 + 3D 引用
        excel95.xls      真 BIFF5（SaveAs 39 = xlExcel9795）
        biff5mix.xls     BIFF5 的引用/区域编码样本
        xor.xls          Excel 95 + 修改密码 => XOR 混淆（4 字节 FILEPASS）
        formatted.ppt    两页 + 演讲者备注 + 图片 + 粗体run

.EXAMPLE
    pwsh -NoProfile -ExecutionPolicy Bypass -File tools/ms_samples.ps1
#>
[CmdletBinding()]
param(
    [string]$OutDir = (Join-Path $PSScriptRoot '..\samples\ms')
)

$ErrorActionPreference = 'Stop'
$office = 'C:\Program Files\Microsoft Office\Root\Office16'
$png = Join-Path $PSScriptRoot '..\samples\pixel.png'
$OutDir = (Resolve-Path -LiteralPath $OutDir -ErrorAction SilentlyContinue)?.Path
if (-not $OutDir) {
    $OutDir = Join-Path $PSScriptRoot '..\samples\ms'
    New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
    $OutDir = (Resolve-Path -LiteralPath $OutDir).Path
}

# ProgID -> CLSID -> LocalServer32；HKCU 覆盖 HKLM，这就是 WPS 抢占的原因。
$clsids = [ordered]@{
    '{000209FF-0000-0000-C000-000000000046}' = "$office\WINWORD.EXE /Automation"     # Word.Application
    '{00024500-0000-0000-C000-000000000046}' = "$office\EXCEL.EXE /automation"       # Excel.Application
    '{91493441-5A91-11CF-8700-00AA0060263B}' = "$office\POWERPNT.EXE /AUTOMATION"    # PowerPoint.Application
}

$saved = @{}
foreach ($clsid in $clsids.Keys) {
    $key = "HKCU:\Software\Classes\CLSID\$clsid\LocalServer32"
    if (Test-Path $key) {
        $saved[$clsid] = (Get-ItemProperty -Path $key).'(default)'
    } else {
        New-Item -Path $key -Force | Out-Null
        $saved[$clsid] = $null
    }
    # A per-user override wins over the machine registration, so delete it
    Remove-Item -Path $key -Force -ErrorAction SilentlyContinue
    Write-Host ("[redirect] {0} -> {1}" -f $clsid, $clsids[$clsid])
}

function Save-Doc($word, $doc, $name, $password) {
    $path = Join-Path $OutDir $name
    # 目标已存在且带密码时，Word 会弹「请输入文件所需的密码」并挂住整个脚本。
    if (Test-Path $path) { Remove-Item $path -Force }
    if ($password) { $doc.Password = $password }
    $doc.SaveAs2($path, 0)   # 0 = wdFormatDocument97
    $doc.Close(0)
    Write-Host "  wrote $name"
}

function Save-Xls($wb, $name, $format, $password) {
    $path = Join-Path $OutDir $name
    if (Test-Path $path) { Remove-Item $path -Force }
    if ($password) { $wb.SaveAs($path, $format, $password) }
    else { $wb.SaveAs($path, $format) }
    $wb.Close(0)
    Write-Host "  wrote $name"
}

try {
    # ---------------- Word ----------------
    # Word remembers every path it failed to open in `Resiliency\DisabledItems` and
    # then hangs on that path forever, even after the file is fixed. Clear the
    # list before each run so a repaired sample can be checked again.
    Remove-Item -Path 'HKCU:\Software\Microsoft\Office\16.0\Word\Resiliency' -Recurse -Force -ErrorAction SilentlyContinue
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    $doc = $word.Documents.Add()
    $doc.Content.Text = "MS-DOC-001`rSecond paragraph, ASCII only.`rUnicode: 中文测试 Ω≈ç√`rEND-MARKER-42`r"
    Save-Doc $word $doc 'hello.doc'

    # 直接给整篇一次性写文本，再按「第 n 段」取格式：先取 Paragraphs.Item(1) 再改它的
    # .Range.Text 会让那个对象失效，格式最后落到别的段上（WPS 容忍、Word 不容忍）。
    $doc = $word.Documents.Add()
    $doc.Content.Text = "Bold red heading`rItalic underline run`rCentred paragraph`r"
    $doc.Paragraphs.Item(1).Range.Font.Bold = $true
    $doc.Paragraphs.Item(1).Range.Font.Size = 18
    $doc.Paragraphs.Item(1).Range.Font.Color = 255
    $doc.Paragraphs.Item(2).Range.Font.Italic = $true
    $doc.Paragraphs.Item(2).Range.Font.Underline = 1
    $doc.Paragraphs.Item(3).Alignment = 1
    $end = $doc.Content
    $end.Collapse(0)
    $table = $doc.Tables.Add($end, 2, 3)
    $table.Cell(1, 1).Range.Text = 'H1'; $table.Cell(1, 2).Range.Text = 'H2'; $table.Cell(1, 3).Range.Text = 'H3'
    $table.Cell(2, 1).Range.Text = 'a';  $table.Cell(2, 2).Range.Text = 'b';  $table.Cell(2, 3).Range.Text = 'c'
    $end = $doc.Content; $end.Collapse(0); $end.InsertAfter("After table.`r")
    $doc.Sections.Item(1).Headers.Item(1).Range.Text = "MS-HEADER-11"
    # Footnotes.Add(Range, Reference, Text)：第二个参数是**脚注标记**（正文里那个上标），
    # 第三个才是脚注正文。只传两个参数会把正文当标记插进主文档。
    $ref = $doc.Paragraphs.Item(1).Range
    $ref.Collapse(1)
    $doc.Footnotes.Add($ref, "1", "MS-FOOTNOTE-33") | Out-Null
    Save-Doc $word $doc 'formatted.doc'

    $doc = $word.Documents.Add()
    $doc.Content.Text = "MS-LOCKED-DOC-777`rSecond secret line.`r"
    Save-Doc $word $doc 'locked.doc' 'secret123'

    $word.Quit()

    # ---------------- Excel ----------------
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $wb = $excel.Workbooks.Add()
    $f = $wb.Worksheets.Item(1); $f.Name = 'F'
    $g = $wb.Worksheets.Add();   $g.Name = 'G'
    $g.Range('A1').Value2 = 7
    $f.Range('A1').Value2 = 10
    $f.Range('A2').Value2 = 20
    $f.Range('A3').Value2 = 1234.5678
    $f.Range('A4').Value2 = 0.25
    $f.Range('A5').Value2 = (Get-Date '2024-03-05')
    $f.Range('A6').Value2 = 45356
    $f.Range('A3').NumberFormat = '0.00'
    $f.Range('A4').NumberFormat = '0.00%'
    $f.Range('A5').NumberFormat = 'yyyy-mm-dd'
    $f.Range('A6').NumberFormat = '$#,##0.00'
    $formulas = @(
        '=A1+A2', '=A1*A2', '=SUM(A1:A2)', '=A1/A2', '=A1^2', '=IF(A1>5,"big","small")',
        '=AVERAGE(A1:A2)', '=MAX(A1:A2)', '=-A1', '=A1&"x"', '=A1>5', '=ROUND(A3,1)',
        '=SUM(A1:A2)*2', '=ABS(A1-A2)', '=A1<>A2', '=SUM(A1,A2)'
    )
    for ($i = 0; $i -lt $formulas.Count; $i++) {
        $f.Range("B$($i + 1)").Formula = $formulas[$i]
    }
    $f.Range('C1').Formula = "='G'!A1"
    $f.Range('C2').Formula = '=SUM(G!A1,3)'
    Save-Xls $wb 'formulas.xls' 56   # 56 = xlExcel8

    # BIFF5：微软 Excel 真的会写 0x0500
    $wb = $excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1); $ws.Name = 'Old'
    $ws.Range('A1').Value2 = 'BIFF5-TEXT'
    $ws.Range('A2').Value2 = 1234.5
    $ws.Range('A3').Formula = '=A2*2'
    $ws.Range('B1').Value2 = 42
    Save-Xls $wb 'excel95.xls' 39    # 39 = xlExcel9795

    $wb = $excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Range('A1').Value2 = 1; $ws.Range('A2').Value2 = 2; $ws.Range('A3').Value2 = 3
    $ws.Range('B1').Formula = '=A2'
    $ws.Range('B2').Formula = '=A2+1'
    $ws.Range('B3').Formula = '=SUM(A1:A3)'
    $ws.Range('B4').Formula = '=A2*2'
    $ws.Range('B5').Formula = '=B1+C1'
    Save-Xls $wb 'biff5mix.xls' 39

    # 修改密码 + Excel 95 => XOR 混淆（FILEPASS 只有 4 字节）
    $wb = $excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1); $ws.Name = 'Xor'
    $ws.Range('A1').Value2 = 'XOR-SECRET'
    $ws.Range('A2').Value2 = 7.5
    Save-Xls $wb 'xor.xls' 39 'secret123'

    $excel.Quit()

    # ---------------- PowerPoint ----------------
    $step = 'ppt:new'
    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.DisplayAlerts = 1   # ppAlertsNone
    $step = 'ppt:add'; $pres = $ppt.Presentations.Add()
    if (Test-Path (Join-Path $OutDir 'formatted.ppt')) { Remove-Item (Join-Path $OutDir 'formatted.ppt') -Force }
    $step = 'ppt:slide1'; $slide = $pres.Slides.Add(1, 1)   # 1 = ppLayoutTitle
    $step = 'ppt:title'; $slide.Shapes.Title.TextFrame.TextRange.Text = 'MS-PPT-001'
    $step = 'ppt:body'; $slide.Shapes.Item(2).TextFrame.TextRange.Text = 'plain then BOLD run'
    $body = $slide.Shapes.Item(2).TextFrame.TextRange
    $step = 'ppt:bold'; $body.Characters(12, 4).Font.Bold = -1
    $step = 'ppt:notes'; $slide.NotesPage.Shapes.Item(2).TextFrame.TextRange.Text = 'MS-SPEAKER-NOTE hidden text'
    $step = 'ppt:slide2'; $slide2 = $pres.Slides.Add(2, 1)
    $step = 'ppt:title2'; $slide2.Shapes.Title.TextFrame.TextRange.Text = 'MS-PPT-002'
    $step = 'ppt:body2'; $slide2.Shapes.Item(2).TextFrame.TextRange.Text = 'second slide body'
    if (Test-Path $png) {
        $step = 'ppt:picture'; $slide.Shapes.AddPicture($png, 0, -1, 40, 300, 60, 60) | Out-Null
    }
    $step = 'ppt:save'; $pres.SaveAs((Join-Path $OutDir 'formatted.ppt'), 1)   # 1 = ppSaveAsPresentation
    $pres.Close()
    $ppt.Quit()
    Write-Host '  wrote formatted.ppt'
}
catch {
    Write-Warning ("failed at {0}: {1}" -f $step, $_)
}
finally {
    foreach ($clsid in $clsids.Keys) {
        $key = "HKCU:\Software\Classes\CLSID\$clsid\LocalServer32"
        if ($null -ne $saved[$clsid]) {
            New-Item -Path $key -Force | Out-Null
            Set-ItemProperty -Path $key -Name '(default)' -Value $saved[$clsid]
        } else {
            Remove-Item -Path $key -Force -ErrorAction SilentlyContinue
        }
        Write-Host "[restore]  $clsid"
    }
    Get-Process WINWORD, EXCEL, POWERPNT -ErrorAction SilentlyContinue |
        Where-Object { $_.MainWindowHandle -eq 0 } |
        Stop-Process -Force -ErrorAction SilentlyContinue
}
