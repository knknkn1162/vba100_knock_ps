function Run-Macro($book) {
    $ws2 = $book.Worksheets("Sheet2")
    $book.Worksheets("Sheet1").Range("A1:C5").Copy()
    $ws2.Range("A1").PasteSpecial($XlPasteType::xlPasteValues)
    $ws2.Range("A1").PasteSpecial($XlPasteType::xlPasteFormats)
    $book.Parent.CutCopyMode = $false
}
