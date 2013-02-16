namespace SpreadSharp

module Records =

    let private recordsRange headers fields worksheet =
        let columnsCount = Array.length headers
        let length = Seq.length fields + 1 |> string
        XlRange.get worksheet "A1" <| Some (string (char (columnsCount + 64)) + length)

    let private displayRecords records recordType range =
        let headers = Utilities.recordFieldsNames recordType
        let fields = Utilities.fieldsArray records
        let array = Array2D.ofArrays headers fields
        XlRange.setValue range array

    let private displayRecords' records recordType worksheet =
        let headers = Utilities.recordFieldsNames recordType
        let fields = Utilities.fieldsArray records
        let array = Array2D.ofArrays headers fields
        let range = recordsRange headers fields worksheet
        XlRange.setValue range array

    /// <summary>Sends the values of a collection of F# records to an Excel range.</summary>
    /// <param name="rcords">The F# records array.</param>
    /// <param name="recordType">The type of the record.</param>         
    /// <param name="range">The range object.</param>         
    let toRange (records : 'T seq) recordType range = displayRecords records recordType range

    /// <summary>Sends the values of a collection of F# records to an Excel worksheet.</summary>
    /// <param name="rcords">The F# records collection.</param>
    /// <param name="recordType">The type of the record.</param>         
    /// <param name="worksheet">The worksheet object.</param>
    let toWorksheet (records : 'T seq) recordType worksheet = displayRecords' records recordType worksheet

    /// <summary>Saves a collection of records in a workbook using the specified file name.</summary>
    /// <param name="rcords">The records collection.</param>
    /// <param name="recordType">The type of the record.</param>         
    /// <param name="filename">The destination file name.</param>
    let saveAs records recordType filename =
        let app = XlApp.start ()
        let wb = XlWorkbook.add app
        let ws = XlWorksheet.byIndex wb 1
        toWorksheet records recordType ws
        XlWorkbook.saveAs filename wb
        XlWorkbook.close wb //None None None
        XlApp.quit app