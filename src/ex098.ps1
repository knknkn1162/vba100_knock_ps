function Get-RegularAddrs($crng, $nrng) {
    $cperson = @{}
    $crng.Cells | %{ $cperson.Add($_.Value(), @($_.Row, $_.Column)) }

    $dir4 = @(@(0,1), @(0, -1), @(1, 0), @(-1, 0))
    $neighdir = @{}
    $crng.Cells |`
        %{ $r=$_; $neighdir.Add($r.Value(), [string[]]@($dir4 | %{$r.Offset($_[0], $_[1]).Value()} ) )}
    
    $intsecparm = @{PassThru = $true; ExcludeDifferent = $true; IncludeEqual = $true }
    $ret = $nrng.Cells |`
        # Check Different row, column
        ?{ @(Compare-Object $cperson[$_.Value()] @($_.Row, $_.Column) @intsecparm -Sync 0).length -eq 0} | ?{
            $r=$_
            [string[]]$nneigh = @($dir4 | %{ @($r.Offset($_[0], $_[1]).Value()) })
            # check different neighbors
            (Compare-Object $neighdir[$r.Value()] $nneigh @intsecparm).length -eq 0
        } |`
        %{$_.Address()}
    return ,$ret
}

function Run-Macro($app, $book) {
    $cws = $book.Worksheets("座席表（現）")
    $nws = $book.Worksheets("座席表（新）")
    $sz = 6
    $crng = $cws.Range("B5").Resize($sz,$sz)
    $nrng = $nws.Range("B5").Resize($sz,$sz)
    $regcells = Get-RegularAddrs $crng $nrng
    Write-Info $regcells.length
    # check irregular cells
    Compare-Object @($nrng.Cells | %{$_.Address()}) $regcells -PassThru |`
        %{ $nws.Range($_).Interior.Color = $xlEnum.XlRgbColor::rgbYellow }
}
