function Run-Macro($app, $book) {
    $app.Windows | %{
        $_.Zoom = 85
        $_.View = $xlEnum.XlWindowView::xlNormalView
    }
    $book.Worksheets |`
        %{$_.PageSetUp.Orientation = $xlEnum.XlPageOrientation::xlLandscape}
    Add-Type -AssemblyName Microsoft.VisualBasic
    $app.Windows |`
        %{$_.SheetViews} |`
        ?{[string]([Microsoft.VisualBasic.Information]::TypeName($_)) -eq "WorksheetView"} | %{
            # goto(reference, scroll)
            $app.GoTo($_.Sheet.Range("A1"), $true)
            $_.DisplayGridlines = $false
        }
}
