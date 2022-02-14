function convertUTF8($str) {
    # In Windows PowerShell, the default encoding is usually Windows-1252, an extension of latin-1, also known as ISO 8859-1.(c.f: https://docs.microsoft.com/en-us/powershell/scripting/dev-cross-plat/vscode/understanding-file-encoding?view=powershell-7.2 )
    [System.Text.Encoding]::UTF8.GetString([System.Text.Encoding]::GetEncoding("ISO-8859-1").GetBytes($resp.content))
}

function cast2d($app, $arr2) {
    return ,($app.WorkSheetFunction.transpose($app.WorkSheetFunction.transpose($arr2)))
}
Import-Module AngleParse

function Run-Macro($app, $book) {
    $ws = $book.Worksheets(1)
    $url = "https://excel-ubara.com/vba100sample/vba100list.html"
    $resp = wget -UseBasicParsing $url
    $content = convertUTF8($resp.content)
    [string[]]$header = $content | Select-HtmlContent "table > thead > tr > th"
    $ws.Range("A1").Resize(1, $header.length) = $header
    $body = 1..5 | %{
        [string[]]$ret = $content | Select-HtmlContent "table > tbody > tr > td:nth-child($_)"
        ,@($ret)
    }
    $row = @($body[0]).length
    $ws.Range("A2").Resize(@($body[0]).length, 5) = $app.WorkSheetFunction.transpose($body)
    [void]$ws.UsedRange.EntireColumn.AutoFit()
}
