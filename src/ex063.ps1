function Run-Macro($app, $book) {
    $ws = $book.Worksheets.Add($book.Worksheets(1))
    # 2nd time: dummy row
    1..2 | %{[void]$book.Worksheets(2).Rows(1).Copy($ws.Cells($_,1))}
    $ws.Name = "マスタ全体"
    2..$book.Worksheets.Count |`
        %{$book.Worksheets($_)} |`
        %{
        [void]$_.Range("A1").CurrentRegion.Offset(1).Copy(
            # requires two or more rows
            $ws.Range("A1").End($xlEnum.xlDirection::xlDown).Offset(1))
        }
    [void]$ws.Rows(1).Delete()
}
