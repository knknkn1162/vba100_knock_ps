function Run-Macro($app, $book) {
    $basename = [System.IO.Path]::GetFileNameWithoutExtension($book.Name)
    $dir = $book.Path + ("/{0}_BACKUP" -f $basename)
    $date = (Get-Date).AddDays(-30) | Get-Date -f "yyyyMMddhhmm"
    $bfile = "{0}_{1}.xlsm" -f $basename, $date
    ls $dir |`
        ?{$bfile -ge $_.Name} |`
        %{rm -fo $_.FullName}
}
