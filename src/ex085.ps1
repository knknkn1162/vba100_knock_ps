function Get-PayPattern($book, $app) {
    $dic = @{}
    $tws = $book.Worksheets("取引先マスタ")
    $trng = $tws.Range("A1").CurrentRegion
    $app.Intersect($trng, $trng.Offset(1).Columns(1)).Cells |`
        %{ $dic.Add($_.Value(), $_.Offset(0,2).Value()) }
    return $dic
}

function Calc-Day($dt, $day) {
    if($day -eq "末") {
        $dt = (Get-Date $dt -Day 1).AddMonths(1).AddDays(-1)
    } else {
        if([int](Get-Date $dt).Day -gt [int]$day) {$dt = $dt.AddMonths(1)}
        $dt = (Get-Date $dt -Day $day)
    }
    return $dt
}

function Get-PayDay($date, $pat, $book, $app) {
    $ws = $book.Worksheets("祝日マスタ")
    $horidays = $ws.Range("A1").CurrentRegion.Columns(1).Offset(1)
    $rng = $book.Worksheets("支払パターン").Range("A1").CurrentRegion
    $info = @($app.Intersect($rng, $rng.Offset(1).Columns(1)).Cells |`
        ?{$_.Value() -eq $pat}
    )
    if($info.length -ne 1) { return "" }
    $ret = Get-Date $date
    $ret = Calc-Day $ret $info[0].Offset(0,1).Value()
    $mon = [int]($info[0].Offset(0,2).Value() -replace "月後", "")
    $ret = (Get-Date $ret.AddMonths($mon) -Day 1)
    $ret = Calc-Day $ret $info[0].Offset(0,3).Value()
    $ret = $app.WorksheetFunction.Workday($ret.AddDays(1), -1, $horidays)
    return $ret
}

function Run-Macro($app, $book) {
    $ws = $book.Worksheets("入金予定")
    $rng = $ws.Range("A1").CurrentRegion
    $dic = Get-PayPattern $book $app
    $rng.Columns(4).NumberFormatLocal = "yyyy/mm/dd(aaa)"
    $app.Intersect($rng, $rng.Offset(1).Columns(1)).Cells |`
        %{$_.Offset(0,3) = (Get-PayDay $_.Offset(0,2).Value() $dic[$_.Value()] $book $app) }
    [void]$rng.Columns(4).AutoFit()
}
