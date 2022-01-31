function Run-Macro($app, $book) {
    $ws = $book.Worksheets(1)
    $ws.Range("B2").CurrentRegion.EntireColumn.FormatConditions.Delete()

    $fc = $app.Intersect($ws.Range("B2").CurrentRegion, $ws.Range("E:E, G:G")).FormatConditions
    # Add (Type, Operator, Formula1, Formula2)
    $conf1 = $fc.Add(
        $xlEnum.XlFormatConditionType::xlCellValue, 
        $xlEnum.XlFormatConditionOperator::xlLess,
        "90%")
    $conf1.Interior.Color = $xlEnum.XlRgbColor::rgbRed
    $conf2 = $fc.Add(
        $xlEnum.XlFormatConditionType::xlCellValue, 
        $xlEnum.XlFormatConditionOperator::xlLess,
        "100%")
    $conf2.Font.Color = $xlEnum.XlRgbColor::rgbRed
}
