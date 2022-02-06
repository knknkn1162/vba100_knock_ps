function Run-Macro($app, $book) {
    $wdPasteMetafilePicture = 3
    $wdExportFormatPDF = 17
    $spath = "{0}/ex079/doc1.docx" -f $book.Path
    $ws = $book.Worksheets(1)
    $rng = $ws.Range("A1").CurrentRegion
    $rng.CopyPicture()
    try {
        $wdapp = New-Object -ComObject Word.Application
        $wdDoc = $wdApp.Documents.Open($spath)
        $wdDoc.Bookmarks("エクセル表").Select()
        $wdApp.Selection.TypeText(("{0}`n{1}`n" -f $book.Name, $ws.Name))
        # PasteSpecial(IconIndex, Link, Placement, DisplayAsIcon, DataType, IconFileName, IconLabel)
        [void]$wdApp.Selection.PasteSpecial($xlnull, $xlnull, $xlnull, $xlnull, $wdPasteMetafilePicture)
        # ExportAsFixedFormat (OutputFileName, ExportFormat ...)
        $wdDoc.ExportAsFixedFormat($spath.Replace(".docx", ".pdf"), $wdExportFormatPDF)
        [void]$wdDoc.Close($false)
    } finally {
        [void]$wdApp.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wdApp) 
    }
    $app.CutCopyMode = $false
}
