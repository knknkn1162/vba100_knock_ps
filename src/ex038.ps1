function cast2d($app, $arr2) {
    return ,($app.WorkSheetFunction.transpose($app.WorkSheetFunction.transpose($arr2)))
}

function Is-Workday($dt, $app, $horidays) {
    #$app.WorksheetFunction.Workday($dt + 1, -1, $horidays) -eq $dt
    $app.WorksheetFunction.Networkdays($dt, $dt, $horidays) -eq 1
}

function Run-Macro($app, $book) {
    $rng = $book.Worksheets("売上").Range("A1").CurrentRegion
    $rng = $app.Intersect($rng.Offset(1), $rng)
    $cols = $rng.Columns.Count
    $horidays = $book.Worksheets("祝日").Range("A1").CurrentRegion.Columns(1)
    $rng.Columns(1).Cells |`
        group {if(Is-Workday $_.Value2() $app $horidays) {"平日"} else {"土日祝"}} | %{
            $arr = $_.group | %{,$_.Resize(1,$cols)}
            $book.Worksheets($_.Name).Range("A2").Resize($arr.Length, $cols) = cast2d $app $arr
    }
}
