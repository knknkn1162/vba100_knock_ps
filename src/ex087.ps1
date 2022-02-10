function Has-Pattern($ws, $pat) {
    try {
        $rng = $ws.Cells.SpecialCells($xlEnum.XlCellType::xlCellTypeFormulas)
    } catch {
        Write-Info "cells not found"
        return $false
    }
    Write-Info $rng.Address()
    # Find(What, After, LookIn, LookAt, ...
    return ($rng.Find($pat,
        $xlnull,
        $xlEnum.XlFindLookIn::xlFormulas,
        $xlEnum.XlLookAt::xlPart) -ne $null)
}

function cast2d($app, $arr2) {
    return ,($app.WorkSheetFunction.transpose($app.WorkSheetFunction.transpose($arr2)))
}

function Run-Macro($app, $book) {
    $wss = $book.Worksheets | ?{$_.Name -ne "相関表"}
    $wss | %{$_.Name = "{0}`t" -f $_.Name}
    $pats = $wss |`
        %{"'{0}'!" -f $_.Name.Replace("'", "''")}
    $arr = $wss |`
        %{ $ws=$_; ,@($pats |`
            %{if(Has-Pattern $ws $_) {"o"} else {""}} )
        }
    0..($arr.length-1) | %{$arr[$_][$_] = ""}
    $book.Worksheets("相関表").Range("C3").Resize($arr.length, $arr.length) = cast2d $app $arr

    $wss | %{$_.Name = $_.Name -replace "`t$", ""}
}
