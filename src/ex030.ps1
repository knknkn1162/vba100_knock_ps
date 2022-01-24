function Run-Macro($app, $book) {
    $tws = $book.Worksheets("名簿")
    $rng = $book.Worksheets("名簿").Range("A1").CurrentRegion.Columns("B")
    $rows2 = [math]::Truncate($rng.Rows.Count()/2)+1
    Write-Info ("rows2: {0}" -f $rows2)
    $ws = $book.Worksheets("名札")
    $cols = 2
    $ws.Range("A1").Resize(2,$cols).Copy()
    2..$rows2 | %{$ws.Cells($_ * 2 - 1, 1).PasteSpecial($xlEnum.XlPasteType::xlPasteFormats)}
    $app.cutCopyMode = $false

    1..$rows2 | %{$_ * $cols} | %{
        $ws.Cells($_ - 1,1).Resize($cols, 2) = $app.WorksheetFunction.transpose($tws.Cells($_ ,2).Resize(2,$cols))}
}
