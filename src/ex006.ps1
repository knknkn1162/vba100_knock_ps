function Run-Macro($app, $book) {
    $rng = $book.Worksheets(1).Range("A1").CurrentRegion
    $rng = $app.Intersect($rng, $rng.Offset(1))
    $rng.Columns(4).NumberFormatLocal = "\#,##0"
    $rng.Columns(1).Cells | `
        ? {[String]$_.Value() -notlike "*-*"} | `
        % {$_.Offset(0,3).FormulaR1C1 =  "=RC[-2] * RC[-1]"}
}
