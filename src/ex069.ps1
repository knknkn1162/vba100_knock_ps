function Find-MergeCells($app, $rng) {
    [string[]]$arr = @()
    $app.FindFormat.MergeCells = $true
    do {
        try {
            # Find (What, After, LookIn, LookAt, SearchOrder, SearchDirection, MatchCase, MatchByte, SearchFormat)
            $nrng = $rng.Worksheet.Cells.Find("", $rng, $xlnull, $xlnull, $xlnull, $xlnull, $xlnull, $xlnull, $true)
        } catch { break }
        if($nrng -eq $null) { break }
        if($nrng.Address() -in $arr) { break }
        $arr += $nrng.Address()
        $rng = $nrng
    } while($true)
    Write-Info $arr
    return $arr
}
function Run-Macro($app, $book) {
    $book.Worksheets | %{
        $ws = $_
        @(Find-MergeCells $app $ws.Range("A1")) |`
            %{$ws.Range($_)} | %{
                $rng, $val = $_.MergeArea, $_.MergeArea(1).Value()
                $_.UnMerge()
                $rng.Value() = $val
        }
    }
}
