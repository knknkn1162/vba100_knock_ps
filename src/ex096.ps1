function Is-Excel($fpath) {
     return ([System.IO.Path]::GetExtension($fpath) -match "xls")
}
  
function Conf-Connection($conn, [string]$accpath, [bool]$type) {
    $conn.Provider = "Microsoft.ACE.OLEDB.12.0"
    $conn.ConnectionString = $accpath
}

function Create-Query($srcpath, [string]$date, [int]$price, [bool]$isExcel) {
    $arr = cat -en UTF8 $srcpath |`
        %{$_ -replace '\$date', $date} |`
        %{$_ -replace '\$price', $price}
    If($isExcel) { $arr = $arr | %{$_ -replace "\[(.*)\]", "[`$1$]"} }
    return ($arr -join "`n")
}

function Write-SQLResult($ws, $rs) {
    [void]$ws.Cells.Clear()
    $arr = 0..($rs.Fields.Count-1) | %{$rs.Fields($_).Name}
    $ws.Range("A1").Resize(1, $arr.length) = $arr
    # 式.CopyFromRecordset (Data, MaxRows, MaxColumns)
    [void]$ws.Range("A2").CopyFromRecordset($rs)
    $ws.Columns("E").NumberFormatLocal = "yyyy/mm/dd"
    $ws.Columns("F:H").NumberFormatLocal = "#,##0"
    [void]$ws.Range("A1").CurrentRegion.EntireColumn.AutoFit()
}

function Run-Macro($app, $book) {
    $adOpenStatic = 3; $adLockOptimistic = 3
    $ws = $book.Worksheets.Add($book.Worksheets(1))
    $ws.Name = "結果"
    $accpath = Join-Path $book.Path "ex096\DB1.accdb"
    $sqlpath = Join-Path $book.Path "../src/ex096.sql"

    $conn = New-Object -ComObject "ADODB.Connection"
    $flag = Is-Excel($accpath)
    try {
        $conn.Provider = "Microsoft.ACE.OLEDB.12.0"
        $conn.ConnectionString = $accpath
        $conn.Open()
        Write-Info ("[Open] adodb state: {0}" -f $conn.State)
        try {
            $sql = (Create-Query $sqlpath "2021/01/01" 1000000 $flag)
            $rs = New-Object -ComObject ADODB.Recordset
            Write-Info "[Execure] sql in $sqlpath"
            $rs.Open($sql, $conn, $adOpenStatic, $adLockOptimistic)
            Write-SQLResult $ws $rs
        } finally {
            $rs.Close()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($rs) 
        }
    } finally {
        $conn.Close()
        Write-Info ("[Close] adodb state: {0}" -f $conn.State)
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($conn) 
    }
}
