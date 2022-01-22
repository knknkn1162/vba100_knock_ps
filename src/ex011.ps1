function Run-Macro($app, $book) {
    try {
        $rng = $book.Worksheets(1).Cells.SpecialCells($XlCellType::xlCellTypeConstants)
    } catch {}
    $rng.Cells | `
        ? {$_.MergeCells()} | `
        % {[void]$_.AddComment("セル結合されています")}
}
