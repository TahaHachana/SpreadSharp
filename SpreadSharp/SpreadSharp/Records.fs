namespace SpreadSharp

module Records =

    open Microsoft.FSharp.Reflection
    open Collections

    let private fields record =
        FSharpValue.GetRecordFields record
        |> Array.map (fun x ->
            match x with
            | :? string as str -> str
            | _ -> x.ToString()
        )

    /// <summary>Sends the values of a collection of F# records to an Excel worksheet.</summary>
    /// <param name="rcords">The F# records collection.</param>
    /// <param name="recordType">The type of the record.</param>         
    /// <param name="worksheet">The worksheet object.</param>
    let toWorksheet worksheet (records: seq<'T>) =
        let headers =
            FSharpType.GetRecordFields typeof<'T>
            |> Array.map (fun x -> x.Name)
        let columnsCount = Array.length headers
        let rangeString = String.concat "" ["A1:"; string (char (columnsCount + 64)) + "1"]
        let rng = XlRange.get worksheet rangeString
        Array.toRange rng headers
        records
        |> Seq.iteri (fun idx x ->
            let idxString = string <| idx + 2
            let rangeString = String.concat "" ["A"; idxString; ":"; string (char (columnsCount + 64)); idxString]
            let rng = XlRange.get worksheet rangeString
            Array.toRange rng <| fields x
        )

    /// <summary>Saves a collection of records in a workbook using the specified file name.</summary>
    /// <param name="rcords">The records collection.</param>
    /// <param name="recordType">The type of the record.</param>         
    /// <param name="filename">The destination file name.</param>
    let saveAs filename records =
        let app = XlApp.start ()
        let wb = XlWorkbook.add app
        let ws = XlWorksheet.byIndex wb 1
        toWorksheet ws records
        XlWorkbook.saveAs wb filename
        XlWorkbook.close wb
        XlApp.quit app