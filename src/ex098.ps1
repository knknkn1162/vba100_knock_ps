function Get-RegularCells($crng, $nrng) {
    $cperson = @{}
    $crng.Cells |`
        %{ $cperson.Add($_.Value(), @($_.Row, $_.Column)) }

    $dir4 = @(@(0,1), @(0, -1), @(1, 0), @(-1, 0))
    $neighdir = @{}
    $intparm = @{PassThru = $true; ExcludeDifferent = $true; IncludeEqual = $true }
    $crng.Cells |`
        %{ $r=$_; $neighdir.Add($r.Value(), [string[]]@($dir4 | %{$r.Offset($_[0], $_[1]).Value()} ) )}
    $ret = $nrng.Cells |`
        # Check Different row, column
        ?{ @(Compare-Object $cperson[$_.Value()] @($_.Row, $_.Column) @intparm -Sync 0).length -eq 0} |`
        ?{
            $r=$_
            [string[]]$nneigh = @($dir4 | %{ @($r.Offset($_[0], $_[1]).Value()) })
            # check different neighbors
            (Compare-Object $neighdir[$r.Value()] $nneigh @intparm).length -eq 0
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
    $regcells = Get-RegularCells $crng $nrng
    Write-Info $regcells.length
    # check irregular cells
    $nrng.Cells |`
        ?{$_.Address() -notin $regcells} |`
        %{ $_.Interior.Color = $xlEnum.XlRgbColor::rgbYellow }
}
