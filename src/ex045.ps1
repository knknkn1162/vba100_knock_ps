function Add-Column($lstobj, [int]$col, [string]$formula, [String]$caption) {
    [void]$lstobj.ListColumns.Add($col)
    $lstobj.HeaderRowRange($col) = $caption
    $lstobj.DataBodyRange.Columns($col) = $formula
}
function Run-Macro($app, $book) {
    $ws = $book.Worksheets(1)
    $lstobj = $ws.Range("B2").ListObject
    Add-Column $lstobj 4 "=sum([@[列1]:[列3]])" "合計列1"
    Add-Column $lstobj 7 "=sum([@[列4]:[列5]])" "合計列2"
}
