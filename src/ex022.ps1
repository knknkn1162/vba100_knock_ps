function fizzbuzz($cols, $mod, $str) {
    $ulimit = 500
    $rng = $ws.Range("1:{0}" -f $ulimit).Columns($cols).Cells |`
        ?{$_.Row % $mod -eq 0}
    [void]$rng.EntireRow.ClearContents()
    $rng.Value() = $str
}

function Run-Macro($app, $book) {
    $ws = $book.Worksheets(1)
    fizzbuzz 1 1 "=Row()"
    fizzbuzz 2 3 "Fizz"
    fizzbuzz 3 5 "Buzz"
    fizzbuzz 4 15 "FizzBuzz"
}
