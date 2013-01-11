namespace SpreadSharp

module internal Array2D =
    
    let ofArray array =
        let length = Array.length array
        Array2D.init length 1 (fun x _ -> array.[x])

    let ofArrays (array : 'T []) array' =
        [|
            yield array
            for x in array' do yield x
        |]
        |> array2D