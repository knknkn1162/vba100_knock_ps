function getFormulaCells($ws) {
    $ret = @()
    try {
        $ret = @($ws.Cells.SpecialCells($xlEnum.XlCellType::xlCellTypeFormulas).Cells)
    } catch {}
    return $ret
}
function Run-Macro($app, $book) {
    $book.Worksheets | %{
        $_.Name = "{0}`t" -f $_.Name
    }
    $book.Worksheets | %{
        $pat = "'{0}'!" -f ($_.Name -replace "'", "''")
        $sht = $_
        #$sht.Cells.SpecialCells($xlEnum.XlCellType::xlCellTypeFormulas).Cells | %{
        getFormulaCells($sht) | %{
            $met = if($_.HasArray()) {"FormulaArray"} else {"Formula2"}
            $prev = $_.Formula2
            $_.$met = $_.Formula2.Replace($pat, "")
            Write-Info ("{0} -> {1}(pat: {2})" -f $prev, $_.$met, $pat)
        }
    }
    $book.Worksheets | %{
        $_.Name = $_.Name -replace "`t$", ""
    }
}
