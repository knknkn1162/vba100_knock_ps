function Run-Macro($app, $book) {
    $names = $book.Names
    $regstr = [regex]::escape('#REF!')
    # keep array even if single element
    $vs = @($names | ?{$_.Visible -eq $false})
    $refs = @($names | ?{$_.RefersTo -match $regstr})
    $vs | %{$_.Visible = $true}
    $refs | %{$_.Delete()}
    # msgbox function is is included in [Microsoft.VisualBasic.Interaction]
    #[Microsoft.VisualBasic.Interaction]::Msgbox(("visible:{0}, delete:{1}" -f $vs.Count, $refs.Count))
    Write-Info ("visible:{0}, delete:{1}" -f $vs.Count, $refs.Count)
}
