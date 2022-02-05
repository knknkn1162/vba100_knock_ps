function Run-Macro($app, $book) {
    $rng = $book.Worksheets(1).Range("A1").CurrentRegion
    $cols = $rng.Columns.Count()
    $cands = $app.Intersect($rng, $rng.Offset(1,1).Columns(1)).Cells | `
        %{,$_.Resize(1,$cols-1)} |`
        ?{($_.Value() | measure -sum).Sum -ge 350} |`
        ?{($_.Value() | ?{$_ -lt 50}).Length -eq 0} |`
        %{$_.Resize(1,1).Offset(0,$cols-1) = "合格"}
}
