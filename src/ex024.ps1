function narrow($str) {
    return [Microsoft.VisualBasic.Strings]::StrConv($str, [Microsoft.VisualBasic.VbStrConv]::Narrow)
}
function conv($str) {
    Add-Type -AssemblyName Microsoft.VisualBasic
    $arr = $str.ToCharArray() |` %{
        $ch = narrow($str)
        If ($ch -match "[A-Z a-z 0-9]") {$ch} else {$_}
    }
    return $arr.join("")
}
function Run-Macro($app, $book) {
    $str = "あいうＡＢＣアイウａｂｃ１２３"
    $str = conv($str)
    Write-Info ("str -> " + $str)
}
