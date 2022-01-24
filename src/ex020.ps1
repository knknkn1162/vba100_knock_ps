function Run-Macro($app, $book) {
    $dir = $book.Path + "/ex020_BACKUP"
    If ( !(Test-Path $dir) ) { mkdir $dir }
    $fname = "ex020_{0}.xlsm" -f (Get-Date -f "yyyyMMddhhmm")
    $book.SaveCopyAs(("{0}/{1}" -f $dir, $fname))
}
