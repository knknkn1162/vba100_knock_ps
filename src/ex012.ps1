function Run-Macro($app, $book) {
    $rng = $book.Worksheets(1).Range("A1").CurrentRegion.Offset(1).Columns(3).Cells |`
        ? {$_.MergeCells()} |`
        ? {$_.MergeArea(1).Address() -eq $_.Address()}
    foreach($r in $rng) {
        $rng2 = $r.MergeArea()
        $val = $r.Value(); $cnt = $rng2.Count
        $rng2.Unmerge()
        $rng2.Value() = [Math]::Truncate($val/$cnt)
        $rng2.Resize($val % $cnt) | %{$_.Value() += 1}
    }
}
