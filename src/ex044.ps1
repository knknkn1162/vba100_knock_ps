function cast2d($app, $arr2) {
    return ,($app.WorkSheetFunction.transpose($app.WorkSheetFunction.transpose($arr2)))
}

function Run-Macro($app, $book) {
    $arr = $book.Worksheets | %{
        $ws=$_; $_.ListObjects | %{
            ,@($_.Name, $ws.Name, $_.DataBodyRange.Address(), $_.DataBodyRange.Rows.Count, $_.DataBodyRange.Columns.Count)
        }
    }
    $tws = $book.Worksheets.Add($xlnull, $book.Worksheets($book.Worksheets.Count))
    $tws.Name = "status"
    $tws.Range("A1").Resize(1,5) = @(
        "テーブル名", "シート名", "セル範囲", "リスト行数", "リスト列数"
    )
    $tws.Range("A2").Resize($arr.Length, 5) = cast2d $app $arr
    [void]$tws.UsedRange.EntireColumn.AutoFit()
}
