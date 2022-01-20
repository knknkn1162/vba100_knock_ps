function Run-Macro($book) {
    $ws2 = $book.Worksheets("Sheet2")
    $book.Worksheets("Sheet1").Range("A1:C5").Copy()
    @($XlPasteType::xlPasteValues, $XlPasteType::xlPasteFormats) | `
        % { $ws2.Range("A1").PasteSpecial($_) }
    $book.Parent.CutCopyMode = $false
}
