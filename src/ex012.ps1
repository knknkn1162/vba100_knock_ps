function Run-Macro($app, $book) {
    $book.Worksheets(1).Range("A1").CurrentRegion.Offset(1).Columns(3).Cells |`
        ? {$_.MergeCells()} |`
        ? {$_.MergeArea(1).Address() -eq $_.Address()} | `
        % {
            $rng2 = $_.MergeArea(); $rng2.Unmerge()
            $val, $cnt = $_.Value(), $rng2.Count
            $rng2.Value() = [Math]::Truncate($val/$cnt)
            $rng2.Resize($val % $cnt) | %{$_.Value() += 1}
        }
}
