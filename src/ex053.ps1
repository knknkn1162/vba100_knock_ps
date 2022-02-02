function Run-Macro($app, $book) {
    $ws = $book.Worksheets(1)
    $prev = (Get-Date "2021/12/31").AddYears(-35)
    $rng = $ws.Range("A1").ListObject.DataBodyRange.Columns(1).Cells |`
        ?{$_.offset(0,1).Value() -eq "男"} |`
        ?{$_.offset(0,2).Value() -le $prev} |`
        ?{$_.offset(0,3).Value() -eq "東京都"} |`
        %{$_.offset(0,4) = "対象"}
}
