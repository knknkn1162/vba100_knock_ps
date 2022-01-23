function Run-Macro($app, $book) {
    try {
        $rng = $book.Worksheets(1).Cells.SpecialCells($xlEnum.XlCellType::xlCellTypeConstants, $xlEnum.XlSpecialCellsValue::xlTextValues)
    } catch {
        Write-Info "not found"
        return
    }
    $rng.Cells | %{
        # "in place" replacement
        $_.Value() = $_.Value() -replace "`r`n", "`n"
        [String[]]$arr = $_.Value() -split "`n" | ?{$_ -ne ""}
        $_.Value() = $arr -join "`n"
        Write-Info ("{0}:{1}" -f $_.Address(), $_.Value())
    }
}
