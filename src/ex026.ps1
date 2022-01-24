function cast2d($app, $arr2) {
    return ,($app.WorkSheetFunction.transpose($app.WorkSheetFunction.transpose($arr2)))
}

function Run-Macro($app, $book) {
    $basename = [System.IO.Path]::GetFileNameWithoutExtension($book.Name)
    $bdir = "{0}/{1}" -f $book.Path, $basename
    $mat = ls $bdir |`
        %{,@($_.Name, $_.LastAccessTime, $_.Length)}
    $ws = $book.Worksheets("ファイル一覧")
    $ws.Range("A1").Resize(1,3) = @("ファイル一覧", "更新日時", "サイズ")
    $ws.Range("A2").Resize($mat.length, 3) = cast2d $app $mat

    $rng = $ws.Range("A1").CurrentRegion
    $app.Intersect($rng, $rng.Columns(1).Offset(1)) |`
        ?{[System.IO.Path]::GetExtension($_.Value()) -like "`.xls*"} |`
        # Add (Anchor, Address, SubAddress, ScreenTip, TextToDisplay)
        %{[void]$ws.Hyperlinks.Add($_, ("{0}/{1}" -f $bdir, $_.Value()), $xlnull, $xlnull, $_.Value()) }

    [void]$ws.UsedRange.EntireColumn.AutoFit()
}
