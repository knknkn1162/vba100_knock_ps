# Open workbooks by hand seems troublesome, so we run automatically.
function Run-BeforeOpenHook($app) {
    $root = Split-Path $PSScriptRoot -parent
    $script_name = Split-Path -Leaf $PSCommandPath
    $basename = [System.IO.Path]::GetFileNameWithoutExtension($script_name)
    $dir = "{0}/books/{1}" -f $root, $basename
    Write-Info("target directory: {0}" -f $dir)
    ls $dir |`
        %{[void]$app.Workbooks.Open($_.FullName)}
    return
}

function Run-Macro($app, $book) {
    $basename = [System.IO.Path]::GetFileNameWithoutExtension($book.Name)
    $txtfile = "{0}/{1}/log_{2}.txt" -f $book.Path, $basename, (Get-Date -f "yyyyMMddhhmmss")
    Write-Info ("logfile: {0}" -f $txtfile)

    $app.Workbooks |`
        ?{$_.FullName -like "*`.xls*"} |`
        # exclude $book itself
        ?{$_.FullName -ne $book.FullName} |`
        %{echo $_.FullName >> $txtfile}
    # After Run-Macro, Application.Quit will called
}
