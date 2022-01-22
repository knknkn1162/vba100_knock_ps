function Run-Macro($app, $book) {
    $ws = $book.Worksheets(1)
    try {
        $rng = $ws.Cells.SpecialCells($xlEnum.XlCellType::xlCellTypeConstants, $xlEnum.XlSpecialCellsValue::xlTextValues)
    } catch {
        Write-Info "cells not found"
        return
    }

    $rng | % {
        $r=$_
        [regex]::Matches($r.Value(), "注意") | `
            # vba is 1-indexed
            %{$r.Characters($_.index+1,2).Font} | `
            # vba color is formatted BBGGRR
            %{$_.Color=$xlEnum.XlRgbColor::rgbRed; $_.Bold = $true}
    }
}
