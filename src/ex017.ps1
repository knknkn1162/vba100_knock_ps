
# TODO: Is There any way to cast jagged array to multi-dimensional array?
function cast2d($app, $arr2) {
    return ,($app.WorkSheetFunction.transpose($app.WorkSheetFunction.transpose($arr2)))
}
function Run-Macro($app, $book) {
    $ws = $book.Worksheets("部・課マスタ")
    $rng = $book.Worksheets("社員").Range("A1").CurrentRegion.Columns("C")
    $rng = $app.Intersect($rng, $rng.Offset(1))
    $arr = $rng.Cells |`
        %{($_.Resize($xlnull,4).Value()) -join ","} |`
        sort |`
        unique |`
        # avoid to flatten 1d-array
        %{,($_ -split ",")}
    Write-Info ("arr type: " + $arr.GetType())

    $ws.Range("A2").Resize($arr.length,4) = cast2d $app $arr
}
