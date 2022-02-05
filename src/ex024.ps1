function Tonarrow($str) {
    return [Microsoft.VisualBasic.Strings]::StrConv($str, [Microsoft.VisualBasic.VbStrConv]::Narrow)
}
function conv($str) {
    Add-Type -AssemblyName Microsoft.VisualBasic
    $arr = $str.ToCharArray() |` %{
        $ch = Tonarrow($_)
        If ($ch -match "[0-9,A-Z,a-z]") {$ch} else {$_}
    }
    return -join $arr
}
function Run-Macro($app, $book) {
    $str = "あいうＡＢＣアイウａｂｃ１２３"
    $str2 = conv($str)
    Write-Info ("before:{0}" -f $str)
    Write-Info ("after :{0}" -f $str2)
}
