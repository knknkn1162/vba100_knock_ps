function cast2d($app, $arr2) {
    return ,($app.WorkSheetFunction.transpose($app.WorkSheetFunction.transpose($arr2)))
}

function GetSheetCounts($sht) {
    [int]$ret = 0
    try {
        $ret = $sht.PageSetup.Pages.Count
    } catch {}
    return $ret
}

function Run-Macro($app, $book) {
    $ws = $book.Worksheets.Add($book.Worksheets(1))
    $ws.Name = "目次"
    $ws.Range("A1").Resize(1,2) = @("シート名", "印刷ページ数")
    $arr = $book.Worksheets |`
        %{,@($_.Name, (GetSheetCounts $_))}
    $ws.Range("A2").Resize($arr.Length, 2) = cast2d $app $arr
    $ws.Range("A2").Resize($arr.Length).Cells |`
        ?{$book.Worksheets($_.Value()).Visible -eq $xlEnum.XlSheetVisibility::xlSheetVisible} |`
        # https://docs.microsoft.com/ja-jp/office/vba/api/excel.hyperlinks.add
        # expression.Add(Anchor, Address, SubAddress, ScreenTip, TextToDisplay)
        %{[void]$ws.Hyperlinks.Add($_, "", "'{0}!A1" -f ($_.Value() -replace "'", "''"))}
}
