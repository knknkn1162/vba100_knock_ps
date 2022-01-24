function Run-Macro($app, $book) {
    $ws = $app.Activesheet.Hyperlinks |`
        ?{$_.Type -eq $msoEnum.MsoHyperlinkType::msoHyperlinkRange} |`
        %{
            $_.Range.Offset(0,1) = $_.Address
            $_.Delete()
    }
}
