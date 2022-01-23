function Set-AppConfig($app, [bool]$arg) {
    $app.Calculation = if($arg){ 
            $xlEnum.xlCalculation::xlCalculationAutomatic
        }else{ 
            $xlEnum.xlCalculation::xlCalculationManual
        }
    $app.DisplayAlerts = $arg
    $app.ScreenUpdating = $arg

}
function Run-Macro($app, $book) {
    $regstr = "*社外秘*"
    Set-AppConfig $app $false
    # change value only & goto A1 cell
    $book.Worksheets |`
        ?{$_.Name -notlike $regstr} |`
        %{
            $_.Visible = $xlEnum.XlSheetVisibility::xlSheetVisible
            [void]$_.Cells.Copy()
            [void]$_.Cells.PasteSpecial($xlEnum.XlPasteType::xlPasteValues)
            $app.CutCopyMode = $false
            # goto(reference, scroll)
            $app.Goto($_.Range("A1"), $true)
        }
    # delete
    $book.Sheets |`
        ?{$_.Name -like $regstr} |`
        %{$_.Delete()}
    
    Set-AppConfig $app $true
}
