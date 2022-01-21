function Run-Macro($app, $book) {
    $ws2 = $book.Worksheets("Sheet2")
    $book.Worksheets("Sheet1").Range("A1:C5").Copy()
    @($XlPasteType::xlPasteValues, $XlPasteType::xlPasteFormats) | `
        % { $ws2.Range("A1").PasteSpecial($_) }
    # TODO: value__ is necessary because $app.cutCopyMode permits only [Microsoft.Office.Interop.Excel.XlCutCopyMode] but $true/$false is also OK
    $app.cutCopyMode.__value = $false
}
