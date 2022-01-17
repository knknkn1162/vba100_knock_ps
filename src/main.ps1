Param(
    [string]$path,
    [bool]$visible = $false
)

# write here
function Run-Macro($book) {
    $ws1 = $book.Worksheets("Sheet1")
    $ws2 = $book.Worksheets("Sheet2")
    $ws1.Range("A1:C5").Copy($ws2.Range("A1"))
    $input = Read-Host "press any key.."
}

try {
    $app = New-Object -ComObject Excel.Application
    $app.visible = $visible
    if $($visible -eq $false) {
        $excel.DisplayAlerts = $false
    }
    try {
        $book = $app.Workbooks.Open($path)
        Run-Macro($book)
    } finally {
        [void]$book.Close($false)
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($book)
    }
} finally {
    [void]$app.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) 
    [GC]::Collect()
}
