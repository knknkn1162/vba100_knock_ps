[int]$ppSaveAsPDF = 32

function Align-Image($old, $new) {
    $new.LockAspectRatio = $true
    $new.Width = $old.Width
　　　　$new.Height = $old.Height
    $new.Width = [math]::min($old.Width, $new.Width)
    $new.Height = [math]::min($old.Height, $new.Height)
    $new.Top = $old.Top + ($old.Top - $new.Top)/2
    $new.Left = $old.Left + ($old.Width - $new.Width)/2
}

function Run-Macro($app, $book) {
    $fdir = "{0}/ex071" -f $book.Path 
    $fpath = "{0}/prezen1.pptx" -f $fdir
    $spath = $fpath.Replace(".pptx", ".pdf")
    if(Test-Path $spath) {rm -r -fo $spath}
    $book.Worksheets(1).ChartObjects(1).Chart.CopyPicture()
    try {
        $pptapp = New-Object -ComObject PowerPoint.Application
        $ppt = $pptApp.Presentations.Open($fpath)
        $pptSlide = $ppt.Slides(1)
        $pptOrigShape = $pptSlide.Shapes(1)
        $pptNewShape = $pptSlide.Shapes.Paste()
        Align-Image $pptOrigShape $pptNewShape
        $app.CutCopyMode = $false
        $pptOrigShape.Delete()
        $ppt.SaveAs($spath, $ppSaveAsPDF)
    } finally {
        [void]$pptApp.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($pptApp) 
    }
}
