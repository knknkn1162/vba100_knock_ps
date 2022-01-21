function Run-Macro($app, $book) {
    $rng = $book.Worksheets(1).Range("B2").CurrentRegion
    $rng = $app.Intersect($rng, $rng.Offset(1))
    $rng.Columns(3).NumberFormatLocal = "\#,##0"
    $rng.Columns(1).Cells | `
        ? {[String]$_.Value() -ne ""} | `
        ? {[String]$_.Offset(0,1).Value() -ne ""} | `
        % {$_.Offset(0,2) = $_.Offset(0,1).Value() * $_.Value();}

}
