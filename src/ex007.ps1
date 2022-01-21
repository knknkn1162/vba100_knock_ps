function Run-Macro($app, $book) {
    $rng = $book.Worksheets(1).Range("A1").CurrentRegion
    $rng = $app.Intersect($rng, $rng.Offset(1).Columns(1))
    $rng.Offset(0,1).NumberFormatLocal = "@"
    [String[]]$arr = $rng.Cells.Value() | `
        %{ $_ -replace " ", "/"} | `
        %{ $_ -replace "元年", "1年" } | `
        Get-Date -f "MMdd"
    $arr.GetType()
    # See http://officetanaka.net/excel/vba/tips/tips124.htm
    $rng.Offset(0,1) = $app.WorksheetFunction.transpose($arr)
}
