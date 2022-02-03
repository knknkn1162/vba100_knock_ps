function zlookup([String]$str, $rng, [int]$idx, [int]$ord) {
    if($ord -ge 1) {$ord--}
    $arr = $rng.Columns(1).Cells | ?{$_.Value() -eq $str} | %{$_.Row}
    Write-Info "match: $arr"
    if($arr.length -eq 0) { return "" }
    try {
        $ret = $rng.Cells($arr[$ord], $idx).Value()
    } catch {$ret = "=na()"}
    return $ret
}
function Run-Macro($app, $book) {
    $ws = $book.Worksheets(1)
    $n1 = zlookup "sample20" $ws.Range("A1").CurrentRegion 2 3
    $n2 = zlookup "sample20" $ws.Range("A1").CurrentRegion 2 0
    $n3 = zlookup "sample50" $ws.Range("A1").CurrentRegion 2 3
    $n4 = zlookup "sample20" $ws.Range("A1").CurrentRegion 2 -1
    $n5 = zlookup "sample20" $ws.Range("A1").CurrentRegion 2 100
    $ws.Range("E1").Resize(1,5).Value() = @($n1,$n2,$n3,$n4,$n5)
}
