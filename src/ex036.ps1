function Run-Macro($app, $book) {
    $ws = $book.Worksheets(1)
    $rng = $ws.Range("A1").CurrentRegion
    $arr = $rng.Rows(1).Cells |`
        %{,@($_.Column, ($_.Value() -match "\(([0-9]+)\)$")) } |`
        %{,@($_[0], $Matches[1])} |`
        sort -Property {[int]$_[1]} |`
        %{$_[0]}
    $cnt = $arr.Length
    1..$cnt |`
        %{[void]$rng.Columns($arr[$_ - 1]).Copy($rng.Columns($cnt + $_))}
    [void]$rng.Resize(1,$cnt).EntireColumn.Delete()
    [void]$ws.UsedRange.EntireColumn.AutoFit()
}
