function cast2d($app, $arr2) {
    return ,($app.WorkSheetFunction.transpose($app.WorkSheetFunction.transpose($arr2)))
}

function Run-Macro($app, $book) {
    $sws = $book.Worksheets("売上")
    $dbws = $book.Worksheets("DB")
    $rows = $sws.Cells($sws.Rows.Count, 1).End($xlEnum.xlDirection::xlUp).Row
    $units = @(1)
    $units += 1..$rows |`
        ?{[string]$sws.Cells($_,1).Value() -eq ""} |`
        %{$_ + 1}
    $arr = $units |`
        %{,$sws.Cells($_, 1).CurrentRegion} |`
        %{
            # this is necessary (WHY?)
            $rng = $sws.Range($_.Address())
            $base = $rng.Resize(1,1)
            Write-Info $rng.Address()
            $app.Intersect($rng, $rng.Offset(2,2)).Cells |`
            ?{$sws.Cells($base.Row + 1, $_.Column).Value() -notmatch "[1-4]Q計"} |`
            %{
                ,@($base.Value(), $base.Offset(0,1).Value(),
                    $sws.Cells($_.Row, 1).Value(), $sws.Cells($_.Row, 2).Value(),
                    $sws.Cells($base.Row + 1, $_.Column).Value(), $_.Value()
                )
            }
        }
    $dbws.Range("A2").Resize($arr.length, 6) = cast2d $app $arr
}
