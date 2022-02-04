function Run-Macro($app, $book) {
    $fws = $book.Worksheets("フォーマット")
    $frng = $fws.Range("A1").CurrentRegion
    $dic = @{}
    $app.Intersect($frng, $frng.Offset(1).Columns(1)).Cells |`
        %{ $dic.Add($_.Value(), @(($_.Offset(0,1).Value() -eq "N"), [int]$_.Offset(0,2).Value())) }
    $ws = $book.Worksheets("data")
    $rng = $ws.Range("A1").CurrentRegion
    $cols = $rng.Columns.Count
    $arr = $app.Intersect($rng, $rng.Offset(1).Columns(1)).Cells |`
        %{,@($_.Resize(1,$cols).Cells | %{
            $conf = $dic[$ws.Cells(1, $_.Column).Value()]
            $member, $pad = if($conf[0]) { "PadLeft", "0" } else { "PadRight", " " }
            ([string]$_.Value()).$member($conf[1], $pad)
        })
    }
    $fdir = "{0}/ex065" -f $book.Path; mkdir -ea 0 $fdir
    $fpath = "{0}/out.txt" -f $fdir
    $arr | %{-join $_} > $fpath
}
