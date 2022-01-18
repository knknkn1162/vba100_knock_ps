Param(
    [string]$pspath,
    [string]$xlpath,
    [switch]$debug
)

."$pspath"

try {
    $app = New-Object -ComObject Excel.Application
    $saveChanges=$true
    if ($debug -eq $true) {
        $app.visible = $true
        $app.DisplayAlerts = $false
        $saveChanges=$false
    }
    try {
        $book = $app.Workbooks.Open($xlpath)
        Run-Macro($book)
        $input = Read-Host "press any key.."
    } finally {
        [void]$book.Close($saveChanges)
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($book)
    }
} finally {
    [void]$app.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) 
    [GC]::Collect()
}
