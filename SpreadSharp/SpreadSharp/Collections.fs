namespace SpreadSharp

open Microsoft.Office.Interop.Excel

module Collections =

//    let private recordsRange headers fields worksheet =
//        let columnsCount = Array.length headers
//        let length = Seq.length fields + 1 |> string
//        XRange.get worksheet "A1" <| Some (string (char (columnsCount + 64)) + length)
//
//    let private displayRecords records recordType range =
//        let headers = Utilities.recordFieldsNames recordType
//        let fields = Utilities.fieldsArray records
//        let array = Array2D.ofArrays headers fields
//        XRange.setValue range array
//
//    let private displayRecords' records recordType worksheet =
//        let headers = Utilities.recordFieldsNames recordType
//        let fields = Utilities.fieldsArray records
//        let array = Array2D.ofArrays headers fields
//        let range = recordsRange headers fields worksheet
//        XRange.setValue range array

    module Array =

        /// <summary>Sends the values of an array to an Excel range.</summary>
        /// <param name="range">The range object.</param>
        /// <param name="array">The array.</param>
        let toRange range array =
            let array2D = Array2D.ofSeq array
            XlRange.setValue range array2D

        /// <summary>Returns the values of an Excel range as an array.</summary>
        /// <param name="range">The range object.</param>
        /// <returns>The range values as an array.</returns>
        let ofRange (range : Range) =
            let value2 = range.Value2 :?> obj [,]
            let length = value2.Length
            [|
                for x in 1 .. length do
                    yield Array2D.get value2 x 1
            |]

        /// <summary>Sends the values of an array to an Excel worksheet.
        /// The values are stored in the first column.</summary>
        /// <param name="worksheet">The worksheet object.</param>
        /// <param name="array">The array.</param>
        let toWorksheet (worksheet : Worksheet) array =
            let arrayLength = Array.length array
            let rangeString = sprintf "A1:A%d" arrayLength
            let range = XlRange.get worksheet rangeString
            toRange range array

    module List =

        /// <summary>Sends the values of a list to an Excel range.</summary>
        /// <param name="range">The range object.</param>
        /// <param name="list">The list.</param>
        let toRange range list =
            let array2D = Array2D.ofSeq list
            XlRange.setValue range array2D

        /// <summary>Returns the values of an Excel range as a list.</summary>
        /// <param name="range">The range object.</param>
        /// <returns>The range values as a list.</returns>
        let ofRange (range : Range) =
            let value2 = range.Value2 :?> obj [,]
            let length = value2.Length
            [
                for x in 1 .. length do
                    yield Array2D.get value2 x 1
            ]

        /// <summary>Sends the values of a list to an Excel worksheet.
        /// The values are stored in the first column.</summary>
        /// <param name="worksheet">The worksheet object.</param>
        /// <param name="list">The list.</param>
        let toWorksheet (worksheet : Worksheet) list =
            let listLength = List.length list
            let rangeString = sprintf "A1:A%d" listLength
            let range = XlRange.get worksheet rangeString
            toRange range list

    module Seq =

        /// <summary>Sends the values of a seq to an Excel range.</summary>
        /// <param name="range">The range object.</param>
        /// <param name="seq">The seq.</param>
        let toRange range seq =
            let array2D = Array2D.ofSeq seq
            XlRange.setValue range array2D

        /// <summary>Returns the values of an Excel range as a seq.</summary>
        /// <param name="range">The range object.</param>
        /// <returns>The range values as a seq.</returns>
        let ofRange (range : Range) =
            let value2 = range.Value2 :?> obj [,]
            let length = value2.Length
            seq {
                for x in 1 .. length do
                    yield Array2D.get value2 x 1
            }

        /// <summary>Sends the values of a seq to an Excel worksheet.
        /// The values are stored in the first column.</summary>
        /// <param name="worksheet">The worksheet object.</param>
        /// <param name="seq">The seq.</param>
        let toWorksheet (worksheet : Worksheet) seq =
            let seqLength = Seq.length seq
            let rangeString = sprintf "A1:A%d" seqLength
            let range = XlRange.get worksheet rangeString
            toRange range seq