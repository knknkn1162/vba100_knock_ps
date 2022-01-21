function Create-WorkSheet($book, [String]$name) {
    # On Error Resume Next
    try {
        [void]$book.Worksheets($name).Delete()
    } catch {
        Write-Warning "[WARNING] Worksheet: ${name} does not exist"
    }
    $cnt = $book.Worksheets.Count()
    $ws = $book.Worksheets.Add($xlnull, $book.Worksheets($cnt))
    $ws.Name = $name
    return $ws
}
function Run-Macro($app, $book) {
    $ws1 = $book.Worksheets("成績表")
    $rng = $ws1.Range("A1").CurrentRegion
    $cols = $rng.Columns.Count()
    [String[]]$arr = 2..($rng.Rows.Count()) | `
        ? { $ws1.Cells($_, $cols).Value() -eq "合格"} | `
        % { $ws1.Cells($_, 1).Value() }

    $ws2 = Create-WorkSheet $book "合格者"
    $ws2.Range("A1").Resize($arr.Length) = $app.WorksheetFunction.transpose($arr)
}
