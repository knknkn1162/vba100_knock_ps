function transpose([Object[,]]$arr, [boolean]$direction) {
    $rows = $arr.GetLength(0)
    $cols = $arr.GetLength(1)
    $ret = New-Object "Object[,]" $cols,$rows
    foreach($i in 1..$cols) {
        foreach($j in 1..$rows) {
            [int]$r = $rows - $j + 1
            [int]$c = $cols - $i + 1
            $ret[[int]($i - 1),[int]($j - 1)] = if($direction) {$arr[$r, $i]} else {$arr[$j, $c]}
        }
    }
    return ,$ret
}

function Run-Macro($app, $book) {
    $ws = $book.Worksheets(1)
    $rng = $ws.Range("A1").CurrentRegion
    # 1-indexed(Object[,])
    $arr2 = $rng.Value()
    $ret = transpose $arr2 $false
    $ws.Range("F1").Resize($rng.Columns.Count, $rng.Rows.Count) = $ret
}
