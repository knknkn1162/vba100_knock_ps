function Date2Str($date) {
    return (Get-Date $date -f "yyyyMMdd")
}
function Run-Macro($app, $book) {
    $bakdir = "{0}/ex057_BACKUP" -f $book.Path
    $dic = @{}
    ls $bakdir |`
        sort -p LastWriteTime |`
        %{$dic[(Date2Str $_.LastWriteTime)] = $_.FullName}
    ls $bakdir |`
        ?{!($dic[(Date2Str $_.LastWriteTime)] -eq $_.FullName)} |`
        %{rm -r $_.FullName}


}
