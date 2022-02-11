function formatCode([string]$str, [int]$depth) {
    $indent = 2
    Write-Info ("{0}{1}" -f (" " * $indent * $depth), $str)
    return ("{0}{1}" -f (" " * $indent * $depth), $str)
}

function formatThTd([string]$tag, $rng, [string]$val, [int]$depth) {
    $str = if($val -eq "") {"&nbsp;"} else {$val}
    $rows = $rng.Rows.Count; $columns = $rng.Columns.Count
    $rowspan = if($rows -ge 2) {" rowspan=""{0}""" -f $rows} else {""}
    $colspan = if($columns -ge 2) {" colspan=""{0}""" -f $columns} else {""}
    $code = "<{0}{1}{2}>{3}</{0}>" -f $tag, $rowspan, $colspan, $str
    return (formatCode $code $depth)
}

function ParseThTd($rng, [string]$tag, [int]$depth) {
    $arr = @()
    $depth++
    $arr += (formatCode "<tr>" $depth)
    $arr += $rng.Cells |`
        ?{$_.MergeArea(1).Address() -eq $_.Address()} |`
        %{formatThTd $tag $_.MergeArea() $_.Value() $depth}
    $depth--
    $arr += (formatCode "</tr>" $depth)
    return $arr
}

function ParseTbody($rng, [int]$depth) {
    $arr = @()
    $arr += (formatCode "<tbody>" $depth)
    $depth++
    $arr += $rng.Columns(1).Cells |`
        %{ (ParseThTd $_.Resize(1, $rng.Columns.Count) "td" $depth) }
    $depth--
    $arr += (formatCode "</tbody>" $depth)
    return $arr
}

function ParseThead($rng, [int]$depth) {
    $arr = @()
    $arr += (formatCode "<thead>" $depth)
    $depth++
    $arr += $rng.Columns(1).Cells |`
        %{ (ParseThTd $_.Resize(1, $rng.Columns.Count) "th" $depth) }
    $depth--
    $arr += (formatCode "</thead>" $depth)
    return $arr
}

function ParseTable($rng, [int]$hnum) {
    $arr = @()
    $depth = 0
    $arr += (formatCode "<table border=""1"">" $depth)
    $depth++
    $arr += (ParseThead $rng.Resize($hnum) $depth)
    $arr += (ParseTbody $app.Intersect($rng, $rng.Offset($nhum)) $depth)
    $depth--
    $arr += (formatCode "</table>" $depth)
    return $arr
}

function ConvertHTML($rng, [int]$hnum) {
    return (ParseTable $rng $hnum) -join "`n"
}
function Run-Macro($app, $book) {
    $sdir = $book.FullName.Replace(".xlsm", ".html")
    $str = (ConvertHTML $book.Worksheets(1).Range("B2").CurrentRegion 2)
    echo $str > $sdir
}
