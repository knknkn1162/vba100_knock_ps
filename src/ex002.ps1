function Run-Macro($app, $book) {
    $ws2 = $book.Worksheets("Sheet2")
    [void]$book.Worksheets("Sheet1").Range("A1:C5").Copy()
    @($xlEnum.XlPasteType::xlPasteValues, $xlEnum.XlPasteType::xlPasteFormats) | `
        % { [void]$ws2.Range("A1").PasteSpecial($_) }
    $app.cutCopyMode = $false
}
