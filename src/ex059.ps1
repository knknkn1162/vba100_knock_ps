function Run-Macro($app, $book) {
    $bdt = [datetime]"2020/04"
    $sdir = "{0}/ex059_out" -f $book.Path
    if(Test-Path $sdir) { rm -r $sdir }
    mkdir -ea 0 $sdir
    0..($book.Sheets.Count-1) |`
        group {[Math]::truncate($_ / 3)} | %{
            $arr = $_.Group |`
                %{$bdt.AddMonths($_)} |`
                %{Get-Date $_ -f "yyyy年MM月"}
            Write-Info ("copy sheets: {0}" -f ($arr -join ","))
            $idx = $arr | %{$book.Worksheets($_).index}
            # $book.Sheets($arr).copy() doesnot work (WHY?)
            $book.Sheets($idx).copy()
            $app.ActiveWorkbook.SaveAs(("{0}/{1}Q.xlsx" -f $sdir, ([int]$_.Name + 1)), $xlEnum.XlFileFormat::xlOpenXMLWorkbook)
            $app.ActiveWorkbook.Close()
        }
    Write-Info ("{0} -> {1}" -f $sdir, ((ls $sdir | %{$_.Name}) -join ","))
}
