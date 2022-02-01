function Add-Columns($lstobj, [int]$col, [string]$formula, [String]$caption) {
    $lstobj.Columns.Add($col)
    $lstobj.HeaderRowRange($col) = $caption
    $lstobj.DataBodyRange.Columns($col) = $formula
}
function Run-Macro($app, $book) {
    $ws = $book.Worksheets(1)
    $lstobj = $ws.Range("B2").ListObject(1)
    Add-Columns $lstobj 4 "=sum([@[列4]:[列5]])" "合計列1"
    Add-Columns $lstobj 7 "=sum([@[列1]:[列3]])" "合計列2"
}
