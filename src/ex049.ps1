function cast2d($app, $arr2) {
    return ,($app.WorkSheetFunction.transpose($app.WorkSheetFunction.transpose($arr2)))
}

function Copy-CondFmtCells($app, $wsIn, $wsOut, $pos, $member, $color) {
    $rng = $wsIn.Range("A1").CurrentRegion
    $cols = $rng.Columns.Count
    $arr = $rng.Columns(1).Cells |`
        ?{$_.Offset(0,3).DisplayFormat.$member.Color -eq $color} |`
        %{,$_.Resize(1,$cols)}
    $wsOut.Cells($pos,1).Resize($arr.Length, $cols) = cast2d $app $arr
    $wsOut.Cells($pos,4).Resize($arr.Length).$member.Color = $color
    return $pos + $arr.Length
}

function Run-Macro($app, $book) {
    $wsIn = $book.Worksheets("49In")
    $wsOut = $book.Worksheets("49Out")
    [void]$wsIn.Range("A1").Resize(1,4).Copy($wsOut.Range("A1"))
    $pos = 2
    $pos = Copy-CondFmtCells $app $wsIn $wsOut $pos "Font" $xlEnum.XlRgbColor::rgbRed
    $pos = Copy-CondFmtCells $app $wsIn $wsOut $pos "Interior" $xlEnum.XlRgbColor::rgbRed
    $pos = Copy-CondFmtCells $app $wsIn $wsOut $pos "Interior" $xlEnum.XlRgbColor::rgbYellow
}
