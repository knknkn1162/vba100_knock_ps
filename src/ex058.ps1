function arr2str([int[]]$arr, [int]$seq) {
    # set guard
    $arr += ($arr[-1] + 2)
    $delta = 1..($arr.Length-1) | %{$arr[$_] - $arr[$_ - 1]}
    $spl = @(0)
    $spl += 1..($delta.length) | ?{$delta[$_ - 1] -ne 1}

    # more than (or equal to) n-sequential number
    $spl2 = 1..($spl.length-1) | %{if( $spl[$_] - $spl[$_-1] -lt $seq){ $spl[$_-1]..$spl[$_]} else {$spl[$_]} }
    # $arr.length-1 : indexer guard
    $spl2 = @(0) + $spl2 | unique
    
    Write-Info "index: $spl2"
    $ret = 1..($spl2.length-1) |`
        %{ (@($arr[$spl2[$_-1]], $arr[$spl2[$_]-1]) | unique) -join "-" }
    return ($ret -join ",")
}
function Run-Macro($app, $book) {
    [int[]]$arr = @(1,2,3,5,8,9,11,12,13,14,15,17,19,20,21,22)
    $ws = $book.Worksheets(1)
    1..5 | %{$ws.Cells($_,1) = arr2str $arr $_}
}
