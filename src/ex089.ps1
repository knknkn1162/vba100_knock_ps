function Run-Macro($app, $book) {
    $dirs = @("ex089_A", "ex089_B", "ex089_C") |`
        %{Join-Path $book.Path $_}
    if(Test-Path $dirs[2]) { rm -r $dirs[2] }
    mkdir -ea 0 $dirs[2]
    # /E: 空のディレクトリを含むサブディレクトリをコピー
    # /XO : 古いファイルを除外
    0..1 | %{robocopy $dirs[$_] $dirs[2] /e /xo}
}
