function normalize($str) {
    $pat = @("（株）", "(株)", "株）", "株)", "（株", "(株", [char]0x3291, [char]0x3231, [char]0x337F, [char]0x33cd) |`
        %{ [regex]::escape($_) }
    $pat = "({0})" -f ($pat -join "|")
    Write-Info $pat
    return ($str -replace $pat, "株式会社")
}
function Run-Macro($app, $book) {
    $book.Worksheets(1).Cells.SpecialCells($xlEnum.XlCellType::xlCellTypeConstants, $xlEnum.XlSpecialCellsValue::xlTextValues).Cells |`
        %{$_.Value() = normalize($_.Value())}
}
