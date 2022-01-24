function Run-Macro($app, $book) {
    $basename = [System.IO.Path]::GetFileNameWithoutExtension($book.Name)
    $bdir = "{0}/{1}" -f $book.Path, $basename
    If (!(Test-Path $bdir)) { rm -r -fo $bdir }
    mkdir -ea 0 $bdir
    $book.Sheets | %{
        $fname = "{0}/{1}.xlsx" -f $bdir, ($_.Name -replace "_", "/")
        $fdir = Split-Path $fname -parent; mkdir -ea 0 $fdir
        Write-Info ("save Sheets({0}) as {1}" -f $_.Name, $fname)
        $_.Copy()
        $app.ActiveWorkbook.SaveAs($fname)
        $app.ActiveWorkbook.Close()
    }
}
