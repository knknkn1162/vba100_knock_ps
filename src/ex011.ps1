function Run-Macro($app, $book) {
    $ws = $book.Worksheets(1)
    try {
        $rng = $ws.Cells.SpecialCells($XlCellType::xlCellTypeConstants)
    } catch {
        Write-Info "merged cells not found. exit."
        return
    }
    [String[]]$arr = $rng.Cells | `
        ? {$_.MergeCells()} | `
        % {$_.MergeArea(1).Address()} | sort | unique
    Write-Info ("AddComment: " + ($arr -join ","))
    $ws.Range($arr -join ",") | `
        %{[void]$_.AddComment("セル結合されています")}
}
