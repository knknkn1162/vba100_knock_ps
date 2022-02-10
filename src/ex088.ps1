function cast2d($app, $arr2) {
    return ,($app.WorkSheetFunction.transpose($app.WorkSheetFunction.transpose($arr2)))
}

function Get-Rank($rate) {
    switch($rate) {
        {$_ -lt 0.5} { return "A" }
        {$_ -lt 0.9} { return "B" }
        default { return "C" }
    }
}

function Calc-ABC($rng, $col) {
    $crng = $rng.Columns($col)
    $sum = ($crng.Value() | measure -sum).Sum
    # Sort (Key1, Order1, Key2, Type, Order2, Key3, Order3, Header...
    [void]$rng.Sort($crng, $xlEnum.XlSortOrder::xlDescending)
    return $crng.Cells |`
        % -b {$rate=0} -p {$rate += $_.Value()/$sum; $rate} |`
        %{Get-Rank($_)}
}

function Run-Macro($app, $book) {
    $dic = @{}
    $book.Worksheets("商品マスタ").Range("A1").CurrentRegion.Offset(1).Columns(1).Cells |`
        ?{[string]$_.value() -ne ""} |`
        %{$dic.Add($_.Value(), @($_.Offset(0,1).Resize(1,3).Value()) )}
    $arr = $book.Worksheets("data").Range("A1").CurrentRegion.Offset(1).Columns(1).Cells |`
        ?{[string]$_.value() -ne ""} |`
        %{
            ,@($_.Value(), $dic[$_.Value()][0],
                $_.Offset(0,1).Value(), $dic[$_.Value()][1],
                $dic[$_.Value()][2])
    }

    $tws = $book.Worksheets("クロスABC")
    # A:E
    $tws.Range("A2").Resize($arr.length, 5) = cast2d $app $arr
    # F:J
    $rng = $tws.Range("A1").CurrentRegion
    $rng = $app.Intersect($rng, $rng.Offset(1))
    $rng.Columns("F").Formula = "=C2*D2"
    $rng.Columns("G").Formula = "=C2*E2"
    $rng.Columns("H").Formula = "=G2-F2"
    $rng.Columns("J").Formula = $app.WorkSheetFunction.transpose((Calc-ABC $rng "H"))
    $rng.Columns("I").Formula = $app.WorkSheetFunction.transpose((Calc-ABC $rng "G"))
}
