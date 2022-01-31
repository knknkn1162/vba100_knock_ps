function Init-Chart($series) {
    $series.Interior.Color = $xlEnum.XlRgbColor::rgbBlue
    $series.ApplyDataLabels($xlEnum.XlDataLabelsType::xlDataLabelsShowNone)
}

function ChartFormat($series, [int]$value, $color) {
    1..($series.Values.Count) |`
        ?{$series.Values.Item($_) -eq $value} |`
        %{
            $tmp=$series.Points($_)
            $tmp.Interior.Color = $color
            $tmp.ApplyDataLabels()
        }
    return
}

function Run-Macro($app, $book) {
    $ws = $book.Worksheets(1)
    $cht = $ws.ChartObjects(1).Chart
    $cht.SetSourceData($ws.Range("A1").CurrentRegion)
    $series = $cht.SeriesCollection(1)
    $sermax = ($series.Values | measure -max).Maximum
    $sermin = ($series.Values | measure -min).Minimum
    Init-Chart $series
    ChartFormat $series $sermin $xlEnum.XlRgbColor::rgbRed
    ChartFormat $series $sermax $xlEnum.XlRgbColor::rgbGreen
}
