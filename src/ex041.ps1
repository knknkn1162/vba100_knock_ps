function randomexpr() {
    $a1 = Get-Random -min 10 -max 99
    $a2 = Get-Random -min 2 -max $a1
    $op = Get-Random -max 3
    return ("{0} {1} {2}" -f $a1, "+-*/"[$op], $a2)
}

function Run-Macro($app, $book) {
    $cnt = 10
    $arr = 1..$cnt |`
        %{randomexpr} |`
        ?{$app.Inputbox("{0} = ?" -f $_) -eq ($_ | iex)}
    $book.Worksheets(1).Range("A1") = "{0}問解けた" -f $arr.length
}
