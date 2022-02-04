function Set-LinkPicture($dst, $src) {
    [void]$src.CopyPicture()
    $ws = $dst.Worksheet
    $ws.Paste()
    $pic = $ws.Shapes($ws.Shapes.Count)
    $pic.LockAspectRatio = $true
    $pic.Width = $dst.Width
    $pic.Height = $dst.Height
    $pic.Width = [math]::min($dst.Width, $pic.Width)
    $pic.Height = [math]::min($dst.Height, $pic.Height)
    $pic.Top = $dst.Top
    $pic.Left = $dst.Left + ($dst.Width - $pic.Width)/2
    $app.CutCopyMode = $false
}

function Run-Macro($app, $book) {
    $ws1 = $book.Worksheets("元表1")
    $ws2 = $book.Worksheets("元表2")
    $ws0 = $book.WorkSheets("まとめ")
    $ws0.Shapes | %{[void]$_.Delete()}
    Set-LinkPicture $ws0.Range("A1:J20") $ws1.Range("A1").CurrentRegion
    Set-LinkPicture $ws0.Range("A21:J40") $ws2.Range("A1").CurrentRegion
    $app.Windows(1).SheetViews("まとめ").DisplayGridLines = $false
}
