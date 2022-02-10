function Is-Overlapped([double]$top1, [double]$bottom1, [double]$top2, [double]$bottom2) {
    return $top1 -le $bottom2 -and $top2 -le $bottom1
}

function Count-OverLappedImages($rng, $isdel) {
    $ws = $rng.Worksheet
    Write-Info ("rng: {0}" -f $rng.Address())
    $ret = @($ws.Shapes |`
        ?{$_.Type -eq $msoEnum.MsoShapeType::msoPicture} | ?{
            $sp=$_
            $arr = @($rng.Areas |`
                ?{(Is-Overlapped $sp.Top ($sp.Top + $sp.Height) $_.Top ($_.Top + $_.Height)) } |`
                ?{(Is-Overlapped $sp.Left ($sp.Left + $sp.Width) $_.Left ($_.Left + $_.Width)) }
            )
            Write-Info ($sp.TopLeftCell.Address(), ($arr | %{$_.Address()}) )
            $arr.length -ne 0
        }
    )
    $ans = $ret.length
    if($isdel) {
        $ret | %{[void]$_.Delete()}
    }
    return $ans
}

function Run-Macro($app, $book) {
    $ws = $book.WorkSheets(1)
    $ws.Range("A3") = Count-OverlappedImages $ws.Range("A1") $true
    $ws.Range("A4") = Count-OverlappedImages $ws.Range("B3:F10") $false
    $ws.Range("A5") = Count-OverlappedImages $ws.Range("B3,C6,E8") $false
    $ws.Range("A6") = Count-OverlappedImages $ws.Range("B5:C7") $true
    $ws.Range("A7") = Count-OverlappedImages $ws.Range("E4") $true
    $ws.Range("A8") = Count-OverlappedImages $ws.Range("B3:F10") $true
}
