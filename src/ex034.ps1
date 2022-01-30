function transpose([Object[][]]$arr, [boolean]$direction) {
    #Write-Info $arr[1].Length
    #Write-Info $arr[1,2].GetType()
    #$arr |`
    #    %{$tmp=$_; $_|%{$_}
    #    }
}
function Run-Macro($app, $book) {
    $arr = @(,@())

    # 1-indexed
    $book.Worksheets(1).Range("A1").CurrentRegion.Cells |`
        %{$arr += ,$_}
    Write-Info $arr[2].Length
    Write-Info $arr.Length
    #transpose $book.Worksheets(1).Range("A1").CurrentRegion.Cells 0
}
