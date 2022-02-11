function FormatColor([int]$color) {
    [byte]$drwR, [byte]$drwG, [byte]$drwB, $null = [BitConverter]::GetBytes($color)
    return "#{0}" -f (-join [Drawing.Color]::FromArgb($drwR, $drwG, $drwB).Name[0..5]).ToUpper()
}

function cast2d($app, $arr2) {
    return ,($app.WorkSheetFunction.transpose($app.WorkSheetFunction.transpose($arr2)))
}

function PrintRGB($rng, [int]$type) {
    $prop = If($type -eq 1) { "Interior" } else { "Font" }
    $cols = $rng.Columns().Count
    $rng.Cells(1,1).$prop.Color
    $arr = $rng.Columns(1).Cells |`
        %{,@( $_.Resize(1,$cols) | %{FormatColor($_.$prop.Color)} ) }
    return $arr
}

function Run-Macro($app, $book) {
    Add-Type -AssemblyName System.Drawing
    $ws = $book.Worksheets(1)
    $ret = (PrintRGB $ws.Range("A1:B3") 1)
    # trush is in $ret[0][0], so we slice it(WHY??)
    $ret = $ret[1..$ret.length]
    $ws.Range("A5") = (cast2d $app $ret)
    $ws.Range("A8:B10") = (cast2d $app $ret)
    $ret = (PrintRGB $ws.Range("A1:B3") 2)

    $ret = $ret[1..$ret.length]
    $ws.Range("A11:B13") = (cast2d $app $ret)
}
