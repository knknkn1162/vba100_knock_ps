Param(
    [string]$pspath,
    [string]$xlpath,
    [switch]$debug
)

."$pspath"

try {

    # [Microsoft.Office.Interop.Excel.ApplicationClass]
    [Microsoft.Office.Interop.Excel.ApplicationClass]$app = New-Object -ComObject Excel.Application
    # get constants such as $xlDirection::xlUp
    # Note) [System.type].GetType() requires AssemblyQualifiedName, but it's messy:(
    $app.GetType().Assembly.GetExportedTypes() | `
        ? {$_.isEnum} | `
        %{ Set-Variable `
            -Name $_.Name `
            -Value ("Microsoft.Office.Interop.Excel.{0}" -f ($_.Name) -as [type]) `
        }
    $saveChanges=$true
    if ($debug -eq $true) {
        $app.visible = $true
        $app.DisplayAlerts = $false
        $saveChanges=$false
    }
    try {
        $book = $app.Workbooks.Open($xlpath)
        Run-Macro $app $book
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
