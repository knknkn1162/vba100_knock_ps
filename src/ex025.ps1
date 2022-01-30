function cast2d($app, $arr2) {
    return ,($app.WorkSheetFunction.transpose($app.WorkSheetFunction.transpose($arr2)))
}

function Run-Macro($app, $book) {
    $ws = $book.Worksheets("売上")
    $rng = $ws.Range("A1").CurrentRegion
    $mat = $app.Intersect($rng, $rng.Offset(1,2)) |`
        %{,@(
            $ws.Cells(
                [Math]::Truncate($_.Row/2) * 2,1).Value(),
                $ws.Cells($_.Row,2).Value(),
                $ws.Cells(1,$_.Column).Value(),
                $_.Value()
        )}
    $db_ws = $book.Worksheets("売上DB")
    $book.Worksheets("売上DB").Range("A2").Resize($mat.length, 4) = cast2d $app $mat
    $db_ws.UsedRange.EntireColumn.AutoFit()
}
