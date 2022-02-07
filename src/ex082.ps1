function cast2d($app, $arr2) {
    return ,($app.WorkSheetFunction.transpose($app.WorkSheetFunction.transpose($arr2)))
}

function Run-Macro($app, $book) {
    $shell = New-Object -COMObject Shell.Application
    $dir = join-path (pwd).Path "books"
    $fld = $shell.Namespace($dir)
    $dic = @{}; 0..400 | %{$dic[$fld.GetDetailsOf($null, $_)] = $_}
    # last author not found, so we use "所有者" instead.
    $header = @("ファイル名", "作成者", "所有者", "作成日時", "更新日時", "前回印刷日", "サイズ")
    $arr = ls -File $dir -Name |`
        %{$fld.ParseName($_)} |`
        %{ $file = $_; ,@($header | %{$fld.GetDetailsOf($file, $dic[$_])}) }
    $ws = $book.Worksheets(1)
    $ws.Range("A1").Resize(1, $header.length) = $header
    $ws.Range("A2").Resize($arr.length, $header.length) = cast2d $app $arr
}

