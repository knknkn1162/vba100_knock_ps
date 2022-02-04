function cast2d($app, $arr2) {
    return ,($app.WorkSheetFunction.transpose($app.WorkSheetFunction.transpose($arr2)))
}

function Run-Macro($app, $book) {
    $dir = "{0}/ex066" -f $book.Path
    $arr = ls -file -r $dir |`
        %{,@($_.FullName, $_.LastWriteTime, $_.length)}

    $ws = $book.Worksheets(1)
    $ws.Range("A1").Resize(1,3) = @("フルパス", "更新日時", "ファイルサイズ")
    $ws.Range("A2").Resize($arr.length, 3) = cast2d $app $arr
    $ws.UsedRange.EntireColumn.AutoFit()
}
