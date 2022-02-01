function Run-Macro($app, $book) {
    $ws = $book.Worksheets(1)
    $arr = @("yyyy/mm/dd", "0", "0.00")
    1..$arr.Length |`
        %{$ws.Columns($_).NumberFormatLocal = $arr[$_ - 1]}
    $bdir = "{0}/ex043_out" -f $book.Path
    mkdir -ea 0 $bdir
    # SaveAs(FileName, FileFormat,...)
    $book.SaveAs("{0}/out.csv" -f $bdir, $xlEnum.XlFileFormat::xlCSV)
}
