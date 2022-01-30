Param(
    [string]$pspath,
    [string]$xlpath,
    [switch]$debug
)

Function Write-Info([String]$msg) {
    If($debug) {
        Write-Verbose $msg -verbose
    } else {
        Write-Verbose $msg
    }
}

function Run-BeforeOpenHook($app) {}
function Run-AfterCloseHook($app) {}

# You can override hook function. See also https://stackoverflow.com/a/38753003
."$pspath"

$xlEnum = New-Object -TypeName PSObject
$msoEnum = New-Object -TypeName PSObject
$xlnull = [System.Reflection.Missing]::Value
try {

    # [Microsoft.Office.Interop.Excel.ApplicationClass]
    [Microsoft.Office.Interop.Excel.ApplicationClass]$app = New-Object -ComObject Excel.Application
    # get constants such as $xlDirection::xlUp
    # Note) [System.type].GetType() requires AssemblyQualifiedName, but it's messy:(
    $app.GetType().Assembly.GetExportedTypes() | `
        ? {$_.isEnum} | `
        %{ $xlEnum | Add-Member `
            -MemberType NoteProperty `
            -Name $_.Name `
            -Value ("Microsoft.Office.Interop.Excel.{0}" -f ($_.Name) -as [type]) `
        }
    # get enumerations such as $msoEnum.msoShapeType::msoTextBox
    [Microsoft.Office.Core.MsoShapeType].Assembly.GetExportedTypes() |`
        ? {$_.isEnum} | `
        %{ $msoEnum | Add-Member `
            -MemberType NoteProperty `
            -Name $_.Name `
            -Value ("Microsoft.Office.Core.{0}" -f ($_.Name) -as [type]) `
        }

    $saveChanges=$true
    if ($debug -eq $true) {
        $app.visible = $true
        $app.DisplayAlerts = $false
        $saveChanges=$false
    }
    Run-BeforeOpenHook($app)
    try {
        $book = $app.Workbooks.Open($xlpath)
        Run-Macro $app $book
        $input = Read-Host "press any key.."
    } finally {
        [void]$book.Close($saveChanges)
        Run-AfterCloseHook($app)
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($book)
    }
} finally {
    [void]$app.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) 
    [GC]::Collect()
}
