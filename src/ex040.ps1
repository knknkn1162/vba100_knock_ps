function Run-Macro($app, $book) {
    $shtName = "2020年12月"
    $ws = $book.Worksheets($shtName)
    $basename = [System.IO.Path]::GetFileNameWithoutExtension($book.Name)
    $dir = "{0}/{1}_data" -f $book.Path, $basename
    ls $dir |`
        %{$app.Workbooks.Open($_.FullName)} |`
        ?{($_.Worksheets | %{$_.Name}) -contains $shtName} |`
        %{$_.Worksheets($shtName)} |`
        %{[void]$_.Range("A1").CurrentRegion.Offset(1).Copy(
            $ws.Range("A1").End($xlEnum.xlDirection::xlDown).Offset(1))}
}
function Run-AfterCloseHook($app) {
    $app.Workbooks |`
        %{[void]$_.Close($false)}
}
