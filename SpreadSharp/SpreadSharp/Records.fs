namespace SpreadSharp

module Records =

    let private recordsRange headers fields worksheet =
        let columnsCount = Array.length headers
        let length = Seq.length fields + 1 |> string
        let rangeString = String.concat "" ["A1:"; string (char (columnsCount + 64)) + length]
        XlRange.get worksheet rangeString

    let private displayRecords (records:seq<'T>) range =
        let headers = Utilities.recordFieldsNames typeof<'T>
        let fields = Utilities.fieldsArray records
        let array = Array2D.ofSeqs headers fields
        XlRange.setValue range array

    let private displayRecords' (records:seq<'T>) worksheet =
        let headers = Utilities.recordFieldsNames typeof<'T>
        let fields = Utilities.fieldsArray records
        let array = Array2D.ofSeqs headers fields
        let range = recordsRange headers fields worksheet
        XlRange.setValue range array

    /// <summary>Sends the values of a collection of F# records to an Excel range.</summary>
    /// <param name="range">The range object.</param>         
    /// <param name="records">The F# records array.</param>
    let toRange range records = displayRecords records range

    /// <summary>Sends the values of a collection of F# records to an Excel worksheet.</summary>
    /// <param name="worksheet">The worksheet object.</param>
    /// <param name="records">The F# records collection.</param>
    let toWorksheet worksheet records = displayRecords' records worksheet

    /// <summary>Saves a collection of records in a workbook using the specified file name.</summary>
    /// <param name="filename">The destination file name.</param>
    /// <param name="records">The records collection.</param>
    let saveAs filename records =
        let app = XlApp.startHidden()
        let wb = XlWorkbook.add app
        let ws = XlWorksheet.byIndex wb 1
        toWorksheet ws records
        XlWorkbook.saveAs wb filename
        XlWorkbook.close wb
        XlApp.quit app