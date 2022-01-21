function Run-Macro($app, $book) {
    $rng = $book.Worksheets(1).Range("A1").CurrentRegion
    $cols = $rng.Columns.Count()
    $app.Intersect($rng, $rng.Offset(1,1).Columns(1)).Cells | `
        ? {($_.Resize($xlnull,$cols-1).Value() | measure -sum).Sum -ge 350} | `
        ? {$app.WorksheetFunction.CountIf($_.Resize($xlnull,$cols-1), 50) -eq 0} | `
        % {$_.Offset($xlnull,$cols-1) = "合格"}
    
}
