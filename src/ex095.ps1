function Get-Text($shape) {
    $ret = ""
    try {
        $ret = $shape.TextFrame.Characters().Text
    } catch {}
    return $ret
}

function cast2d($app, $arr2) {
    return ,($app.WorkSheetFunction.transpose($app.WorkSheetFunction.transpose($arr2)))
}

function Run-Macro($app, $book) {
    $pat = $app.Inputbox("検索文字列を入力してください")
    $delstr = "[{0}]" -f $book.Name.Replace("'", "''")
    $arr = $book.Worksheets |`
        %{$_.Shapes} |`
        ?{(Get-Text $_) -match $pat} | %{
            ,@(
                # Range.Address(RowAbsolute, ColumnAbsolute, ReferenceStyle, External, RelativeTo)
                $_.TopLeftCell.Address($xlnull, $xlnull, $xlnull, $true).Replace($delstr,""),
                (Get-Text $_)
            )
        }
    try {
        $book.Worksheets("検索結果").Delete()
    } catch {}
    $ws = $book.Worksheets.Add($book.Worksheets(1))
    $ws.Name = "検索結果"
    $ws.Range("A1").Resize(1,2) = @("セルアドレス", "図形テキスト")
    $ws.Range("A2").Resize($arr.length, 2) = cast2d $app $arr

    # add hyperlinks
    $rng = $ws.Range("A1").CurrentRegion.Columns(1)
    $app.Intersect($rng, $rng.Offset(1)).Cells |`
        # expression.Add(Anchor, Address, SubAddress, ScreenTip, TextToDisplay)
        %{ [void]$ws.Hyperlinks.Add($_, "", $_.Value().Replace("'", "''"), $xlnull, $_.Value()) }
    [void]$ws.UsedRange.EntireColumn.AutoFit()
}
