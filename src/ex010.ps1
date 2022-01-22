function Run-Macro($app, $book) {
    $rng = $book.Worksheets(1).Range("A1").CurrentRegion.Offset(1).Columns(4).Cells | `
        ? { [String]$_.Offset(0,-1).Value() -eq "" } | `
        ? { $_.Value() -match "[(削除)|(不要)]" }
    Write-Info ("Delete: " + ($rng.Cells | %{$_.EntireRow.Address()}))
    [void]$rng.EntireRow.Delete()
}
