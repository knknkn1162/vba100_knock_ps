function Run-Macro($app, $book) {
    $dic = @{}
    $rng = $book.Worksheets("ID").Range("A1").CurrentRegion
    $app.Intersect($rng, $rng.Offset(1)).Columns(1).Cells |`
        %{ $dic.Add($_.Value(), $_.Offset(0,1).Value()) }
    foreach($i in 1..3) {
        $id, $pass = [string]$app.Inputbox("IDを入力"), [string]$app.Inputbox("passを入力")
        Write-Info ($id, $pass)
        if($dic[$id] -eq $pass) {Write-Info "validated"; return }
    }
    Write-Info "login failed"
    $orig = $app.DisplayAlerts; $app.DisplayAlerts = $false
    If($app.Workbooks.Count -ge 2) { $book.Close($false) }
    $app.DisplayAlerts = $orig
}
