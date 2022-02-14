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

function rotate([ref]$arrref) {
    $arr = $arrref.Value
    $col = $arr.getLength(1) - 1
    $row = $arr.getLength(0) - 1
    $move = Get-Random -min 1 -max $row
    $step = Get-Random -min 1 -max $row
    $rotnum = 0..($row-1) | %{$move + $_ * $step}
    Write-Info "rotate: $rotnum"
    1..$row | %{
        $r=$_
        $brr = @(1..$col | %{ $arr[$r, [int](($_ + $rotnum[$r-1]) % $col + 1)] })
        0..($col-1) | %{$arr[$r, [int]($_ + 1)] = $brr[$_]}
    }
}

function shuffle($row, $col, $app) {
    $arr2 = New-Object "System.Object[,]" ($row+1), ($col+1)
    0..($row-1) |`
        %{$r=$_; 0..($col-1) |`
            %{ $arr2[$r, $_] = $r * $col + $_ } 
        }
    $arr2 = $app.WorksheetFunction.transpose($arr2)
    rotate([ref]$arr2)
    $arr2 = $app.WorksheetFunction.transpose($arr2)
    rotate([ref]$arr2)
    $arr = @(1..$row | %{$r=$_; 1..$col | %{$arr2[$r, $_]}})
    Write-Info $arr
    return $arr
}

function Run-Macro($app, $book) {
    $cws = $book.Worksheets("座席表（現）")
    $nws = $book.Worksheets("座席表（新）")
    $sz = 6
    $crng = $cws.Range("B5").Resize($sz,$sz)
    $nrng = $nws.Range("B5").Resize($sz,$sz)
    $num = $sz*$sz
    $carr = $crng.Cells | %{$_.Value()}
    $retry = 0
    do {
        [void]$nrng.ClearContents()
        $shuf = shuffle $sz $sz $app
        $nrng.Cells | % -b {$i=0} -p {$_.Value() = $carr[$shuf[$i]]; $i++ }
        $addrs = Get-RegularAddrs $crng $nrng
        Write-Info $addrs.length
        if($addrs.length -eq $num) { break }
    } while($true)
}
