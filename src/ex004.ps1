function Run-Macro($app, $book) {
    $rng = $book.Worksheets(1).Range("A1").CurrentRegion
    $app.Intersect($rng.Offset(1,1), $rng).SpecialCells($xlEnum.XlCellType::xlCellTypeConstants).ClearContents()
}
