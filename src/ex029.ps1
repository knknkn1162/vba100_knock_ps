function Run-Macro($app, $book) {
    $fd = $app.FileDialog($msoEnum.MsoFileDialogType::msoFileDialogFilePicker)
    $fd.Filters.Clear()
    [void]$fd.Filters.Add("画像ファイル", "*.jpg;*.png;*.bmp;*.jpeg;*.gif")
    $fd.InitialFileName = $book.Path
    $fd.AllowMultiSelect = $false
    $path = ""
    Write-Info "select file."
    If ($fd.Show() -eq -1) {
        $path = $fd.SelectedItems(1)
    }
    Write-Info ("path: {0}" -f $path)
    $rng = $app.ActiveCell

    # AddPicture (FileName, LinkToFile, SaveWithDocument, Left, Top, Width, Height)
    $shp = $app.ActiveSheet.Shapes.AddPicture(
        $path, $false, $true, $rng.Left, $rng.Top, -1, -1
    )
    $shp.Width = $app.WorksheetFunction.min($rng.Width, $shp.Width)
    $shp.Height = $app.WorksheetFunction.min($rng.Height, $shp.Height)
    $shp.Left = $shp.Left + ($rng.Width-$shp.Width)/2
    $shp.Top = $shp.Top + ($rng.Height-$shp.Height)/2
}
