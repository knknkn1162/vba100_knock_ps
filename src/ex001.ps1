function Run-Macro($app, $book) {
    $ws1 = $book.Worksheets("Sheet1")
    $ws2 = $book.Worksheets("Sheet2")
    $ws1.Range("A1:C5").Copy($ws2.Range("A1"))
}
