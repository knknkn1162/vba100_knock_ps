function Create-Table($rng, $arr, $app) {
    $sz = $arr.length
    $w = $rng.Resize($sz+1, $sz+1)
    [void]$w.Clear()
    $rng.Offset(1).Resize($sz) = $app.WorksheetFunction.transpose($arr)
    $rng.Offset(0,1).Resize(1,$sz) = $arr
    $w.Borders.LineStyle = $xlEnum.XlLineStyle::xlContinuous
    $w.HorizontalAlignment = $xlEnum.Constants::xlCenter
    1..$sz |`
        %{$rng.Offset($_, $_)} |`
        %{$_.Borders($xlEnum.XlBordersIndex::xlDiagonalDown).LineStyle = $xlEnum.XlLineStyle::xlContinuous}
}

function Run-Macro($app, $book) {
    try {
        $app.Workbooks.Delete("相関表")
    } catch {}
    $arr = $book.Worksheets | %{$_.Name}
    $ws = $book.WorkSheets.Add($book.Worksheets(1)); $ws.Name = "相関表"
    Create-Table $ws.Range("B2") $arr $app

}
