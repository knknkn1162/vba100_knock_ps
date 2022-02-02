function Run-Macro($app, $book) {
    $arr = $book.Sheets |`
        ?{$_.Visible -eq $xlEnum.XlSheetVisibility::xlSheetVisible} |`
        ?{$_.Name -like "*印刷*"} |`
        %{$_.Name}
    # 式.PrintOut (From, To, Copies, Preview)
    $book.Sheets($arr).PrintOut($xlnull, $xlnull, $xlnull, $true)
}
