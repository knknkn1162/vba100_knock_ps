function getSheets($app, [String]$path) {
    $ws = $app.Workbooks.Open($path)
    $arr1 = $ws.Sheets |` %{$_.Name}
    [void]$ws.Close($false)
    return $arr1
}
function Run-Macro($app, $book) {
    $basename = [System.IO.Path]::GetFileNameWithoutExtension($book.Name)
    $bdir = "{0}/{1}" -f $book.Path, $basename
    [String[]]$arr1 = getSheets $app ("{0}/Book_{1}.xlsx" -f $bdir, "20201101")
    [String[]]$arr2 = getSheets $app ("{0}/Book_{1}.xlsx" -f $bdir, "20201102")
    Write-Info ("{0} vs {1}" -f ($arr1 -join ","), ($arr2 -join ","))
    $str = if((compare $arr1 $arr2).Length -eq 0) {"一致"}else{"不一致"}
    Write-Info ("シート比較: {0}" -f $str)
}
