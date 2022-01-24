function Run-Macro($app, $book) {
    $checked = "_checked"
    $book.Worksheets(1).Shapes |`
        ?{$_.Type -ne $msoEnum.MsoShapeType::msoFormControl} |`
        ?{$_.Type -ne $msoEnum.MsoShapeType::msoOLEControlObject} |`
        ?{$_.Name -ne $checked} |`
        %{
            $shp = $_.Duplicate()
            $shp.Left = $_.Left + $_.Width
            $shp.Top = $_.Top
            $_.Name = $shp.Name = $checked
    }
}
