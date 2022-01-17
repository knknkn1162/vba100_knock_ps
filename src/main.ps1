Param(
    [string]$pspath,
    [string]$xlpath,
    [bool]$visible = $false
)

."$pspath"

try {
    $app = New-Object -ComObject Excel.Application
    $app.visible = $visible
    if ($visible -eq $false) {
        $app.DisplayAlerts = $false
    }
    try {
        $book = $app.Workbooks.Open($xlpath)
        Run-Macro($book)
        $input = Read-Host "press any key.."
    } finally {
        [void]$book.Close($false)
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($book)
    }
} finally {
    [void]$app.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) 
    [GC]::Collect()
}
