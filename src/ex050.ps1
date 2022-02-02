function Run-Macro($app, $book) {
    [bigint[]]$arr = @()
    $cnt = 1000
    $ret = 3..$cnt | % -b {$arr += @(0,1,1) } -p { $arr += ($arr[-3..-1] | measure -sum).Sum } -e {$arr | %{"'{0}" -f [string]$_}}
    $book.Worksheets(1).Range("A1").Resize($arr.length) = $app.WorksheetFunction.transpose($ret)
    $book.Worksheets(1).UsedRange.EntireColumn.AutoFit()
}
