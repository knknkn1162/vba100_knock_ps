function cast2d($app, $arr2) {
    return ,($app.WorkSheetFunction.transpose($app.WorkSheetFunction.transpose($arr2)))
}

function Run-Macro($app, $book) {
    $sws = $book.Worksheets("売上")
    $rws = $book.Worksheets("請求書")
    $mws = $book.Worksheets("取引先マスタ")
    $rng = $sws.Range("A1").CurrentRegion
    $dic = @{}
    $sdir = "{0}/ex083" -f $book.Path
    If(Test-Path $sdir) { rm -r $sdir }; mkdir -ea 0 $sdir
    $mws.Range("A1").CurrentRegion.Columns(1).Cells |`
        %{$dic.Add($_.Value(), $_.Offset(0,1).Resize(1,4))}
    $app.Intersect($rng, $rng.Offset(1).Columns(1)).Cells |`
        group {$_.Value()} | %{
            [void]$rws.Range("A2:A5,A10:D24").ClearContents()
            $spath = ("{0}/{1}_{2}.pdf" -f $sdir, $_.Name, (Get-Date -f "yyyyMM"))
            $rws.Range("A2:A5") = $app.WorksheetFunction.Transpose($dic[$_.Name])
            $arr1 = $_.group | %{$_.Offset(0,2)}
            $rws.Range("A10:A24").Resize($arr1.length, 1) = $app.WorksheetFunction.Transpose($arr1)
            $arr2 = $_.group | %{,@($_.Offset(0,3).Resize(1,2))}
            $rws.Range("C10:D24").Resize($arr2.length,2) = cast2d $app $arr2
            # 式.ExportAsFixedFormat ( Type , FileName....
            $rws.ExportAsFixedFormat($xlEnum.XlFixedFormatType::xlTypePDF, $spath)
        }
}
