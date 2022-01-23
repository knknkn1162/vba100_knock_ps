function Run-Macro($app, $book) {
    $st = Get-Date "2020/04"
    $cnt = $book.Worksheets.Count()
    1..12 |`
        %{$st.AddMonths($_ - 1) | Get-Date -f "yyyy年MM月"} |`
        %{$book.Worksheets($_).Move($xlnull, $book.Worksheets($cnt))}
}
