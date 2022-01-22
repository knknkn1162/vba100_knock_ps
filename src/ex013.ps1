function Run-Macro($app, $book) {
    $book.Cells. `
    SpecialCells($XlCellType::xlCellTypeConstants, $XlCellType::xlTextValues).Cells | `
    %{$_.Value() -match "注意" }
}
