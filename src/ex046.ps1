function Run-Macro($app, $book) {
    $ws = $book.Worksheets(1)
    [void]$ws.Rows(1).Insert()
    [void]$ws.Rows(2).Copy($ws.Rows(1))
    # 式.CreateNames (Top, Left, Bottom, Right)
    [void]$ws.Range("A1").CurrentRegion.Resize(2).CreateNames($True)
    [void]$ws.Rows(1).Delete()
    $book.Names | %{Write-Info ("name: {0}, address: {1}" -f $_.Name, $_.RefersTo)}
}
