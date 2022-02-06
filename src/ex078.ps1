function countChartObjects($ws) {
    foreach($i in 1..30) {
        try { [void]$ws.ChartObjects($i) } catch { return ($i-1) }
    }
}
function Run-Macro($app, $book) {
    $ws = $book.Worksheets("Sheet1")
    # ChartObjects.Count doesn't work...
    $cnt = countChartObjects($ws)
    Write-Info ("# of charts: {0} vs {1}" -f $ws.ChartObjects.Count, $cnt)
    1..$cnt | %{
        $ser = $ws.ChartObjects($_).Chart.SeriesCollection(1)
        $arr = $ser.Formula -split ","
        Write-Info @($_, $ser.Formula)
        -3..-2 | %{
                $base = $app.Range($arr[$_]).Resize(1,1)
                # Address (RowAbsolute, ColumnAbsolute, ReferenceStyle, External, RelativeTo)
                $arr[$_] = $app.Range($base, $base.End($xlEnum.xlDirection::xlDown)).Address(
                    $xlnull, $xlnull, $xlnull, $true
                )
        }
        $ser.Formula = $arr -join ","
    }
}
