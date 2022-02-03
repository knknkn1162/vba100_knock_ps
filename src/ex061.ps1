function Run-Macro($app, $book) {
    $mrng = $book.Worksheets("マスタ").Range("A1").CurrentRegion
    $dic = @{}
    $app.Intersect($mrng, $mrng.Offset(1)).Columns(1).Cells |`
        %{$dic.Add($_.Value(), $_.Phonetic.Text)}
    $trng = $book.Worksheets("data").Range("A1").CurrentRegion
    $app.Intersect($trng, $trng.Offset(1)).Columns(1).Cells | %{
        if($dic.containsKey($_.Value())) {
            $_.Phonetic.Text = $dic[$_.Value()]
        } else {
            $_.Font.Color = $xlEnum.XlRgbColor::rgbRed
        }
    }
}
