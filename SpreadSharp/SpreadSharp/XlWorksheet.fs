namespace SpreadSharp

open Microsoft.Office.Interop.Excel

module XlWorksheet =

    /// <summary>Adds a worksheet to the specified workbook and optionally sets its name.</summary>
    /// <param name="workbook">The workbook object.</param>
    /// <param name="nameOption">The name of the worksheet.</param>
    /// <returns>The new worksheet.</returns>
    let add (workbook : Workbook) nameOption =
        workbook.Worksheets.Add()
        :?> Worksheet
        |> Com.pushComObj
        |> Utilities.setWorksheetName nameOption

    /// <summary>Adds a worksheet to the specified workbook before another one and optionally sets its name.</summary>
    /// <param name="workbook">The workbook object.</param>
    /// <param name="worksheet">The worksheet before which the new one must be added.</param>
    /// <param name="nameOption">The name of the worksheet.</param>
    /// <returns>The new worksheet.</returns>
    let addBefore (workbook : Workbook) worksheet nameOption =
        workbook.Worksheets.Add(Before = worksheet)
        :?> Worksheet
        |> Com.pushComObj
        |> Utilities.setWorksheetName nameOption

    /// <summary>Adds a worksheet to the specified workbook after another one and optionally sets its name.</summary>
    /// <param name="workbook">The workbook object.</param>
    /// <param name="worksheet">The worksheet after which the new one must be added.</param>
    /// <param name="nameOption">The name of the worksheet.</param>
    /// <returns>The new worksheet.</returns>
    let addAfter (workbook : Workbook) worksheet nameOption =
        workbook.Worksheets.Add(After = worksheet)
        :?> Worksheet
        |> Com.pushComObj
        |> Utilities.setWorksheetName nameOption

    /// <summary>Adds multiple worksheets to the specified workbook.</summary>
    /// <param name="workbook">The workbook object.</param>
    /// <param name="count">The number of worksheets to add.</param>
    let addMany (workbook : Workbook) count =
        [1 .. count]
        |> List.iter (fun _ ->
            workbook.Worksheets.Add()
            :?> Worksheet
            |> Com.pushComObj
            |> ignore)

    /// <summary>Adds multiple worksheets before another one to the specified workbook.</summary>
    /// <param name="workbook">The workbook object.</param>
    /// <param name="worksheet">The worksheet before which the new ones must be added.</param>
    /// <param name="count">The number of worksheets to add.</param>
    let addManyBefore (workbook : Workbook) worksheet count =
        [1 .. count]
        |> List.iter (fun _ ->
            workbook.Worksheets.Add(Before = worksheet)
            :?> Worksheet
            |> Com.pushComObj
            |> ignore)

    /// <summary>Adds multiple worksheets after another one to the specified workbook.</summary>
    /// <param name="workbook">The workbook object.</param>
    /// <param name="worksheet">The worksheet after which the new ones must be added.</param>
    /// <param name="count">The number of worksheets to add.</param>
    let addManyAfter (workbook : Workbook) worksheet count =
        [1 .. count]
        |> List.iter (fun _ ->
            workbook.Worksheets.Add(After = worksheet)
            :?> Worksheet
            |> Com.pushComObj
            |> ignore)

    /// <summary>Moves a worksheet to a new workbook.</summary>
    /// <param name="worksheet">The worksheet to move.</param>
    let move (worksheet : Worksheet) = worksheet.Move()

    /// <summary>Moves a worksheet before another.</summary>
    /// <param name="worksheet">The worksheet to move.</param>
    /// <param name="beforeWorksheet">The worksheet before which the moved one must be placed.</param>
    let moveBefore (worksheet : Worksheet) beforeWorksheet = worksheet.Move(Before = beforeWorksheet)

    /// <summary>Moves a worksheet after another.</summary>
    /// <param name="worksheet">The worksheet to move.</param>
    /// <param name="beforeWorksheet">The worksheet after which the moved one must be placed.</param>
    let moveAfter (worksheet : Worksheet) afterWorksheet = worksheet.Move(After = afterWorksheet)

    /// <summary>Copies a worksheet to a new workbook.</summary>
    /// <param name="worksheet">The worksheet to copy.</param>
    let copy (worksheet : Worksheet) = worksheet.Copy()

    /// <summary>Copies a worksheet before another.</summary>
    /// <param name="worksheet">The worksheet to copy.</param>
    /// <param name="beforeWorksheet">The worksheet before which the copied one must be placed.</param>
    let copyBefore (worksheet : Worksheet) beforeWorksheet = worksheet.Copy(Before = beforeWorksheet)

    /// <summary>Copies a worksheet after another.</summary>
    /// <param name="worksheet">The worksheet to copy.</param>
    /// <param name="afterWorksheet">The worksheet after which the copied one must be placed.</param>
    let copyAfter (worksheet : Worksheet) afterWorksheet = worksheet.Copy(After = afterWorksheet)

    /// <summary>Hides a worksheet.</summary>
    /// <param name="worksheet">The worksheet to hide.</param>
    let hide (worksheet : Worksheet) = worksheet.Visible <- XlSheetVisibility.xlSheetHidden

    /// <summary>Displays a hidden worksheet.</summary>
    /// <param name="worksheet">The hidden worksheet.</param>
    let unhide (worksheet : Worksheet) = worksheet.Visible <- XlSheetVisibility.xlSheetVisible

    /// <summary>Returns the worksheet located at the specified index.</summary>
    /// <param name="workbook">The workbook containing the target worksheet.</param>
    /// <param name="idx">The worksheet index, count starts at 1.<param>
    /// <returns>The target worksheet.</returns>
    let byIndex (workbook : Workbook) (idx : int) = workbook.Worksheets.[idx] :?> Worksheet

    /// <summary>Returns the worksheet having the specified name.</summary>
    /// <param name="workbook">The workbook containing the target worksheet.</param>
    /// <param name="name">The worksheet name.<param>
    /// <returns>The target worksheet.</returns>
    let byName (workbook : Workbook) (name : string) = workbook.Worksheets.[name] :?> Worksheet

    /// <summary>Activates a worksheet.</summary>
    /// <param name="worksheet">The worksheet to activate.</param>
    let activate (worksheet : Worksheet) = worksheet.Activate()

    /// <summary>Activates the worksheet located at the specified index.</summary>
    /// <param name="workbook">The workbook containing the target worksheet.</param>
    /// <param name="idx">The worksheet index, count starts at 1.<param>
    let activateByIndex (workbook : Workbook) idx =
        byIndex workbook idx
        |> activate

    /// <summary>Activates the worksheet having the specified name.</summary>
    /// <param name="workbook">The workbook containing the target worksheet.</param>
    /// <param name="idx">The worksheet index, count starts at 1.<param>
    let activateByName (workbook : Workbook) name =
        byName workbook name
        |> activate