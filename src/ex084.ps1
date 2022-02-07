function Run-Macro($app, $book) {
    $maxcnt = 30
    $bdir = "{0}/ex084_backup" -f $book.Path
    mkdir -ea 0 $bdir
    $fname = $book.Name.Replace(
        ".xlsm",
        "_{0}.xlsm" -f (Get-Date -f "yyMMddhhmmss")
    )
    $spath = "{0}/{1}" -f $bdir, $fname
    Write-Info "saveas $spath"
    $book.SaveCopyAs($spath)
    if(@(ls $bdir).length -le $maxcnt) { return }
    $dpath = (ls $bdir | ?{$_.Name -match "^ex084_"} | sort -p Name | Select -first 1).FullName
    Write-Info "delete $dpath"
    rm $dpath
}
