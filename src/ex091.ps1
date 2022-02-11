function cast2d($app, $arr2) {
    return ,($app.WorkSheetFunction.transpose($app.WorkSheetFunction.transpose($arr2)))
}

function Calc-WorkingHours([string]$start, [string]$end) {
    $stime = If([timespan]$start -lt [timespan]"9:00:00") {[timespan]"9:00:00"} else {$start}
    $hh, $mm, $ss = $end.split(":")
    $etime = New-TimeSpan -hour $hh -min $mm -sec $ss
    $ret = ($etime - $stime) - (New-TimeSpan -hour 9)
    if($ret -le (New-TimeSpan)) {return (New-TimeSpan)} else { return $ret }
}

function Run-Macro($app, $book) {
    $kws = $book.Worksheets("勤怠")
    $rng = $kws.Range("A1").CurrentRegion
    $arr = $app.Intersect($rng, $rng.Offset(1).Columns(1)).Cells |`
        group {@($_.Value(), (Get-Date $_.Offset(0,1).Value() -f "yyyyMM")) -join ","} | %{
            $totalms = ($_.Group |`
                %{ Calc-WorkingHours $_.Offset(0,2).Text() $_.Offset(0,3).Text() } |`
                %{ [math]::truncate($_.TotalMinutes) } |`
                measure -sum
            ).Sum
            Write-Info $totalms
            ,@(
                $_.Name.Split(",")
                ([math]::truncate($totalms / 30) * 30 / (60*24))
            )
        }
    $tws = $book.Worksheets("残業")
    $tws.Range("A2").Resize($arr.length, 3) = cast2d $app $arr
}
