namespace SpreadSharp

open Microsoft.Office.Interop.Excel
open XRange

module Collections =

    module internal Array2D =
    
        let ofArray array =
            let length = Array.length array
            Array2D.init length 1 (fun x _ -> array.[x])

    module Array =

        /// <summary>Sends the values of an array to an Excel range.</summary>
        /// <param name="range">The range object.</param>
        /// <param name="array">The array.</param>
        let toRange range array =
            let array2D = Array2D.ofArray array
            XRange.setValue range array2D

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

    module List =

        /// <summary>Sends the values of a list to an Excel range.</summary>
        /// <param name="range">The range object.</param>
        /// <param name="list">The list.</param>
        let toRange range list =
            let array2D = Array2D.ofArray <| List.toArray list
            XRange.setValue range array2D

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

    module Seq =

        /// <summary>Sends the values of a seq to an Excel range.</summary>
        /// <param name="range">The range object.</param>
        /// <param name="seq">The seq.</param>
        let toRange range seq =
            let array2D = Array2D.ofArray <| Seq.toArray seq
            XRange.setValue range array2D

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