function Separate-BranchData($book, $ws, [string]$key) {
    $ws.AutoFilterMode = $false
    # AutoFilter (Field, Criteria1, Operator, Criteria2, SubField, VisibleDropDown)
    [void]$ws.Range("A1").AutoFilter(1, $key)
    $ws2 = $book.Worksheets($key);
    [void]$ws.Range("A1").CurrentRegion.Copy($ws2.Range("A1"))
    $ws.AutoFilterMode = $false
    return
}
function Save-Worksheet($app, $ws, [string]$fpath) {
    Write-Info $fpath
    $ws.Move()
    $app.ActiveWorkbook.SaveAs($fpath, $xlEnum.XlFileFormat::xlOpenXMLWorkbook)
    $app.ActiveWorkbook.Close()
    return
}

function Run-Macro($app, $book) {
    $ws = $book.Worksheets(1)
    $glob = "{0}/ex093/月別/*.xls*" -f $book.Path
    $sdir = "{0}/ex093/支店別" -f $book.Path
    if(Test-Path $sdir) { rm -r $sdir }
    mkdir -ea 0 $sdir
    $header = @()
    ls -File $glob |`
        %{$app.Workbooks.Open($_.FullName)} | %{
            $rng = $_.Worksheets(1).Range("A1").CurrentRegion
            $header = $rng.Resize(1,$rng.Columns.Count).Value()
            [void]$rng.Offset(1).Copy(
                $ws.Cells($ws.Rows.Count, 1).End($xlEnum.xlDirection::xlUp).Offset(1)
            )
        }
    $app.Workbooks |`
        ?{$_.Name -ne $book.Name} |`
        %{$_.Close($false)}

    $rng = $ws.Range("A2").CurrentRegion
    $rng.EntireColumn.AutoFit()
    $cols = $rng.Columns.Count
    $ws.Range("A1").Resize(1, $header.length) = $header
    # Sort (Key1, Order1, Key2, Type, Order2, Key3, Order3, Header...
    [void]$rng.Sort($ws.Range("A2"), $xlEnum.XlSortOrder::xlAscending)
    $branches = $rng.Columns(1).Value() | unique
    $branches |`
        %{$tmp=$book.Worksheets.Add($book.Worksheets(1)); $tmp.Name = $_}
    $branches | %{Separate-BranchData $book $ws $_}
    $branches |`
        %{$book.Worksheets($_)} |`
        %{Save-WorkSheet $app $_ ("{0}/{1}.xlsx" -f $sdir, $_.Name)}
}
