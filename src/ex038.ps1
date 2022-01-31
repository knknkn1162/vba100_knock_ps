function cast2d($app, $arr2) {
    return ,($app.WorkSheetFunction.transpose($app.WorkSheetFunction.transpose($arr2)))
}

function Is-Workday($dt, $app, $horidays) {
    $app.WorksheetFunction.Workday($dt + 1, -1, $horidays) -eq $dt
}

function Run-Macro($app, $book) {
    $rng = $book.Worksheets("売上").Range("A1").CurrentRegion
    $rng = $app.Intersect($rng.Offset(1), $rng)
    $cols = $rng.Columns.Count
    $horidays = $book.Worksheets("祝日").Range("A1").CurrentRegion.Columns(1)
    $arr = $rng.Columns(1).Cells |`
        # Value2: get serial value
        ?{Is-Workday $_.Value2() $app $horidays} |`
        %{,$_.Resize(1, $cols).Value()}
    $book.Worksheets("平日").Range("A2").Resize($arr.Length, $cols) = cast2d $app $arr

    $arr = $rng.Columns(1).Cells |`
        ?{!(Is-Workday $_.Value2() $app $horidays)} |`
        %{,$_.Resize(1, $cols).Value()}
    $book.Worksheets("土日祝").Range("A2").Resize($arr.Length, $cols) = $arr
}
