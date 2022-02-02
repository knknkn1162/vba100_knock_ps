function Date2Str($date) {
    return (Get-Date $date -f "yyyyMMdd")
}
function Run-Macro($app, $book) {
    $bakdir = "{0}/ex057_BACKUP" -f $book.Path
    $dic = @{}
    ls $bakdir | %{
        $datestr = Date2Str($_.LastWriteTime) 
        $nval = @($_.LastWriteTime, $_.FullName)
        if(!$dic.Contains($datestr)) {
            $dic[$datestr] = $nval
        } else {
            # if $_ is old
            if($dic[$datestr][0] -lt $_.LastWriteTime) {
                $rmname = $dic[$datestr][1]
                $dic[$datestr] = $nval
            } else {
                $rmname = $nval[1]
            }
            rm -r $rmname
        }
    }
}
