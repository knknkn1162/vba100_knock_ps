function Run-Macro($app, $book) {
    $rng = $book.Worksheets(1).Range("A1").CurrentRegion
    # see https://teratail.com/questions/172999
    $app.Intersect($rng, $rng.Offset(1,1)).ClearContents()
}
