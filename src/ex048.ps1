function Is-Numeric($num) {
    return $num.GetType() -in @([int], [double], [decimal], [single], [long])
}
function conv($num) {
    if(Is-Numeric($num)){return [Math]::truncate($num)} else {return $num}
}
function convert2([ref]$arrref) {
    $arr = $arrref.Value
    foreach($i in 0..($arr.GetLength(0)-1)) {
        foreach($j in 0..($arr.GetLength(1)-1)) {
            $arr[$i, $j] = conv($arr[$i, $j])
        }
    }
}
function convert1([ref]$arrref) {
    $arr = $arrref.Value
    foreach($i in 0..($arr.GetLength(0)-1)) {
        $arr[$i] = conv($arr[$i])
    }
}
function convert([ref]$arrref) {
    $arr = $arrref.Value
    if([int]$arr.Rank -eq 0) { return}
    if($arr.Rank -ge 3) { return}
    If ($arr.Rank -eq 1) {
        convert1($arrref)
    } else {
        convert2($arrref)
    }
}

function Run-Macro($app, $book) {
    $arr0 = 3.5
    $arr1 = @(-1.5, 1.5, "1.5", "2020/1/1")
    $arr2 = New-Object "object[,]" 2,4
    0..3 | %{$arr2[0, $_] = $arr1[$_]}
    0..3 | %{$arr2[1, $_] = $arr1[$_]}
    convert([ref]$arr0)
    convert([ref]$arr1)
    convert([ref]$arr2)

    $ws = $book.Worksheets(1)
    $ws.Range("A1") = $arr0
    $ws.Range("A4").Resize(1,4).Value() = $arr1
    $ws.Range("A6").Resize(2,4).Value() = $arr2
}
