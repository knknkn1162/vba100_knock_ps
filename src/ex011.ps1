function Run-Macro($app, $book) {
    $ws = $book.Worksheets(1)
    try {
        $rng = $ws.Cells.SpecialCells($XlCellType::xlCellTypeConstants)
    } catch {}
    [String[]]$arr = $rng.Cells | `
        ? {$_.MergeCells()} | `
        % {$_.MergeArea(1).Address()} | sort | unique
    echo ($arr -join ",")
    $ws.Range($arr -join ",") | `
        %{$_.AddComment("セル結合されています")}
}
