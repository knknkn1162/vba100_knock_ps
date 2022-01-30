function Run-Macro($app, $book) {
    $rng = $book.Worksheets("データ").Range("A1").CurrentRegion
    $rng = $app.Intersect($rng, $rng.Offset(1))
    $rng.Columns("D") = "=IFERROR(VLOOKUP(B2,マスタ!A:C,2,FALSE),"""")"
    $rng.Columns("E") = "=IFERROR(VLOOKUP(B2,マスタ!A:C,3,FALSE),"""")"
    $rng.Columns("F") = "=C2*E2"
    $rng.Columns("D:F").Copy()
    $rng.Columns("D:F").PasteSpecial($xlEnum.XlPasteType::xlPasteValues)
    $app.CutCopyMode = $false
}
