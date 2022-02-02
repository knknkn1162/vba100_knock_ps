function Run-Macro($app, $book) {
    $arr = @()
    $arr += [bigint[]]@(0,1,1)
    $cnt = 1000
    3..$cnt | %{ $arr += ($arr[$_ - 3] + $arr[$_ - 2] + $arr[$_ - 1]) }
    $arr = $arr | %{"'{0}" -f [string]$_}
    $book.Worksheets(1).Range("A1").Resize($arr.length) = $app.WorksheetFunction.transpose($arr)
    $book.Worksheets(1).UsedRange.EntireColumn.AutoFit()
}
