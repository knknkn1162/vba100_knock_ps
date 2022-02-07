function Run-Macro($app, $book) {
    $ws = $book.Worksheets(1)
    $orig = $ws.Selection
    [void]$ws.Cells($ws.Rows.Count, $ws.Columns.Count).Select()
    If($ws.AutoFilterMode) { $ws.AutoFilter.ShowAllData() }
    $arr = $ws.ListObjects |`
        ?{$_.AutoFilter -ne $null} |`
        %{$_.AutoFilter.ShowAllData() }
    [void]$ws.Range("A1").Select()
}
