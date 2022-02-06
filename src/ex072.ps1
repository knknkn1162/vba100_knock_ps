function Run-Macro($app, $book) {
    $ret = @("ＩＴ","と IT","itは","IT 99", "ＧＩＴ","site","It's","it is", "  it  ", "a    it  ", "あ    it  ") |`
        %{$_ -replace "(^|[^A-ZＡ-Ｚ　 ’'])([　 ’']*)([IＩ][TＴ])(?![　 ’']*[A-ZＡ-Ｚ])", "`$1`$2DX"}
    $ret | %{Write-Info $_}
}
