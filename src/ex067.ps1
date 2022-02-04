function cast2d($app, $arr2) {
    return ,($app.WorkSheetFunction.transpose($app.WorkSheetFunction.transpose($arr2)))
}

function Run-Macro($app, $book) {
    $ws = $book.Worksheets(1)
    $rng = $ws.Range("A1").CurrentRegion
    # for test
    $ws.Range("E1") = "東京都"
    $cols = $rng.Columns.Count
    [void]$ws.Columns("F").Resize($cols).Clear()
    $arr = $rng.Columns(1).Cells |`
        ?{$_.Value() -in @("都道府県", $ws.Range("E1").Value())} |`
        %{,$_.Offset(0,1).Resize(1,$cols-1)}
    $ws.Range("F1").Resize($arr.length, $cols-1) = cast2d $app $arr
    # copy-paste formats
    [void]$ws.Columns(2).Resize($xlnull, $cols-1).Copy()
    [void]$ws.Columns(6).Resize($xlnull, $cols-1).PasteSpecial($xlEnum.XlPasteType::xlPasteFormats)
    $app.cutCopyMode = $false
    [void]$ws.UsedRange.EntireColumn.AutoFit()
}
