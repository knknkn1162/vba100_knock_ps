function Run-Macro($app, $book) {
    $rng = $book.Worksheets(1).Range("A1").CurrentRegion
    $rngA = $rng.Columns(1)
    $rngB = $rng.Columns(2)
    $ddd = $rngA.Cells | %{$_.Value}
    ,$ddd | gm
    [int[]]$arr = ($rngA.Value() + $rngB.Value()) | sort | unique
    $rng.Columns(3).Resize($arr.Length, 1) = $app.WorksheetFunction.transpose($arr)
}
