function Run-Macro($app, $book) {
    $rng = $app.ActiveSheet.Range("A1")
    $cond = ($book.Sheets | %{$_.Name}) -join ","
    Write-Info $cond
    $rng.Validation.Delete()
    # Add (Type, AlertStyle, Operator, Formula1, Formula2)
    $validation = $rng.Validation
    $validation.Add(
        $xlEnum.XlDVType::xlValidateList,
        $xlEnum.XlDVAlertStyle::xlValidAlertStop,
        $xlnull,
        $cond,
        $xlnull
    )
    $validation.ErrorTitle = "エラー発生"
    $validation.ShowError = $true
    $validation.ErrorMessage = "シート名が無効です"

}
