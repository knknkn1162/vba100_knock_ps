function Run-Macro($app, $book) {
    $rng = $book.Worksheets(1).Range("A1").CurrentRegion
    $rng.Row(1).Cells |`
        %{$_.Value() -match "\(([0-9]+)\)^"} |`
        %{$Matches}
}
