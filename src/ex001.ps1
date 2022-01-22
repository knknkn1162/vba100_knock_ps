function Run-Macro($app, $book) {
    $book.Worksheets("Sheet1").Range("A1:C5").Copy($book.Worksheets("Sheet2").Range("A1"))
}
