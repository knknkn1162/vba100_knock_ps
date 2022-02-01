
function cast2d($app, $arr2) {
    return ,($app.WorkSheetFunction.transpose($app.WorkSheetFunction.transpose($arr2)))
}

function Run-Macro($app, $book) {
    $ws = $book.Worksheets("階層")
    $rng = $ws.Range("A1").CurrentRegion
    $cols = $rng.Columns.Count
    $arr = New-Object 'Object[]' $cols
    $ws2 = $book.Worksheets("階層DB")
    $ret = $rng.Columns(1).Cells | %{
        $idx = ($_.Resize(1,$cols).Cells |%{[String]$_.Value() -ne ""}).IndexOf($true)
        $arr[$idx] = $_.Offset(0,$idx).Value()
        if($idx -eq 3) {,$arr.Clone()}
    }
    $rng.Resize(1,$cols).Copy($ws2.Range("A1"))
    $ws2.Range("A2").Resize($ret.Length, $cols) = cast2d $app $ret
}
