namespace SpreadSharp

module internal Array2D =
    
    let ofSeq seq = array2D [|seq|]

    let ofSeqs seq seq' =
        [|
            yield  seq
            yield! seq'
        |]
        |> array2D