namespace SpreadSharp

open Microsoft.Office.Interop.Excel

module XRange =

    /// <summary>Returns a worksheet range.</summary>
    /// <param name="worksheet">The worksheet containing the range.</param>
    /// <param name="cell">The first cell in the range.</param>
    /// <param name="cell'">The last cell in the range.</param>
    /// <returns>The range object.</returns>
    let get (worksheet : Worksheet) (cell : string) cell' =
        let cell'' = Utilities.boxOrMissing<string> cell'
        worksheet.Range(cell, cell'')
        |> COM.pushComObj

    /// <summary>Selects a range of cells.</summary>
    /// <param name="range">The range to select.</param>
    let select (range : Range) = range.Select() |> ignore

    /// <summary>Copies a range to the clipboard.</summary>
    /// <param name="range">The range to copy.</param>
    let copy (range : Range) = range.Copy() |> ignore

    /// <summary>Copies a range to the spedified destination.</summary>
    /// <param name="range">The range to copy.</param>
    /// <param name="destination">The destination range.</param>
    let copyToRange (range : Range) (destination : Range) = range.Copy destination |> ignore

    /// <summary>Cuts a range to the clipboard.</summary>
    /// <param name="range">The range to cut.</param>
    let cut (range : Range) = range.Cut() |> ignore

    /// <summary>Cuts a range and pastes it into the specified destination.</summary>
    /// <param name="range">The range to cut.</param>
    /// <param name="destination">The paste destination range.</param>
    let cutPaste (range : Range) (destination : Range) = range.Cut(destination) |> ignore
    
    /// <summary>Deletes a range.</summary>
    /// <param name="range">The range to delete.</param>
    let delete (range : Range) shift =
        let shift' = Utilities.boxOrMissing<XlInsertShiftDirection> shift
        range.Delete shift' |> ignore

    /// <summary>Inserts a cell or a range of cells using the optional
    /// shift direction and copy origin parameters.</summary>
    /// <param name="range">The range representing the column.</param>
    /// <param name="shift">The column index, count starts at 1.</param>
    /// <param name="copyOrigin">The column index, count starts at 1.</param>
    let insert (range : Range) shift copyOrigin =
        let shift' = Utilities.boxOrMissing<XlInsertShiftDirection> shift
        let copy' = Utilities.boxOrMissing<obj> copyOrigin
        range.Insert(shift', copy') |> ignore

    /// <summary>Performs an autofill from a source range to a destination one. The two ranges must overlap.</summary>
    /// <param name="range">The range to copy.</param>
    /// <param name="destination">The destination range.</param>
    /// <param name="autoFillType">The fill type.</param>
    let autoFill (range : Range) destination autoFillType =
        range.AutoFill(destination, autoFillType) |> ignore

    module Column =

        /// <summary>Returns the range representing the column with the specified index.</summary>
        /// <param name="worksheet">The worksheet containing the column.</param>
        /// <param name="idx">The column index, count starts at 1.</param>
        let byIndex (worksheet : Worksheet) (idx : int) =
            worksheet.Columns.[idx]
            :?> Range
            |> fun x -> x.EntireColumn
            |> COM.pushComObj

        /// <summary>Returns the range representing the column with the specified index.</summary>
        /// <param name="worksheet">The worksheet containing the column.</param>
        /// <param name="idx">The column index, count starts at 1.</param>
        let byHeader (worksheet : Worksheet) (header : string) =
            worksheet.Columns.[header]
            :?> Range
            |> fun x -> x.EntireColumn
            |> COM.pushComObj

        /// <summary>Insert a column using the optional shift direction and copy origin parameters.</summary>
        /// <param name="range">The range representing the column.</param>
        /// <param name="shift">The column index, count starts at 1.</param>
        /// <param name="copyOrigin">The column index, count starts at 1.</param>
        let insert (range : Range) shift copyOrigin =
            let shift' = Utilities.boxOrMissing<XlInsertShiftDirection> shift
            let copy' = Utilities.boxOrMissing<obj> copyOrigin
            range.EntireColumn.Insert(shift', copy') |> ignore

        /// <summary>Hides a column.</summary>
        /// <param name="range">The range representing the column to hide.</param>
        let hide (range : Range) = range.EntireColumn.Hidden <- true

        /// <summary>Unhides a column.</summary>
        /// <param name="range">The range representing the hidden column.</param>
        let unhide (range : Range) = range.EntireColumn.Hidden <- false