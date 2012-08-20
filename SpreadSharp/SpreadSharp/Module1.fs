#if INTERACTIVE
#r "Microsoft.Office.Interop.Excel"
#endif

namespace ExcelOp

open System
open Microsoft.Office.Interop.Excel

[<AutoOpenAttribute>]
module Utilities =

    let boxOrMissing = function Some x -> box x | None -> Type.Missing

[<AutoOpenAttribute>]
module Types =

    type TextFileOrigin = Platform of XlPlatform | CodePage of int

    type OpenTextParameters =
        {
            Filename             : string
            Origin               : TextFileOrigin option
            StartRow             : int option
            DataType             : XlTextParsingType option
            TextQualifier        : XlTextQualifier
            ConsecutiveDelimiter : bool option
            Tab                  : bool option
            Semicolon            : bool option
            Comma                : bool option
            Space                : bool option
            Other                : bool option
            OtherChar            : bool option
            FieldInfo            : int [,] option
            TextVisualLayout     : obj option
            DecimalSeparator     : obj option
            ThousandsSeparator   : obj option
            TrailingMinusNumbers : bool option
            Local                : bool option
        }

    type CloseWorkbookParameters =
        {
            SaveChanges : bool option
            Filename : string option
            RouteWorkbook : bool option
        }

module Excel =

    /// Starts Excel with the specified visibility setting.
    let start isVisible = ApplicationClass(Visible = isVisible)

    /// Closes Excel.
    let quit (excel : ApplicationClass) = excel.Quit()


module Workbook =

    /// Adds a workbook to an Excel application.
    let addWorkbook (excel : ApplicationClass) = excel.Workbooks.Add()

    /// Opens an existing workbook.
    let openWorkbook (excel : ApplicationClass) fileName = excel.Workbooks.Open fileName

    /// Closes a workbook.
    let closeWorkbook (workbook : Workbook) (options : CloseWorkbookParameters option) =
        options |> function
            | None -> workbook.Close()
            | Some options' ->
                workbook.Close(
                    SaveChanges   = boxOrMissing options'.SaveChanges,
                    Filename      = boxOrMissing options'.Filename,
                    RouteWorkbook = boxOrMissing options'.RouteWorkbook)

    /// Saves changes to the specified workbook.
    let saveWorkbook (workbook : Workbook) = workbook.Save()

    /// Saves changes to a workbook in a different file.
    let saveWorkbookAs (workbook : Workbook) fileName = workbook.SaveAs fileName

    /// Opens a text file as a new workbook with a single sheet.
    let openText (application : ApplicationClass) (parameters : OpenTextParameters) =
        application.Workbooks.OpenText(
            Filename = parameters.Filename,
            Origin = (parameters.Origin |> function
                | Some x -> x |> function Platform p -> box p | CodePage cp -> box cp
                | None -> Type.Missing),
            StartRow = boxOrMissing parameters.StartRow,
            DataType = boxOrMissing parameters.DataType,
            TextQualifier = parameters.TextQualifier,
            ConsecutiveDelimiter = boxOrMissing parameters.ConsecutiveDelimiter,
            Tab = boxOrMissing parameters.Tab,
            Semicolon = boxOrMissing parameters.Semicolon,
            Comma = boxOrMissing parameters.Comma,
            Space = boxOrMissing parameters.Space,
            Other = boxOrMissing parameters.Other,
            OtherChar = parameters.OtherChar,
            FieldInfo = boxOrMissing parameters.FieldInfo,
            TextVisualLayout = boxOrMissing parameters.TextVisualLayout,
            DecimalSeparator = boxOrMissing parameters.DecimalSeparator,
            ThousandsSeparator = boxOrMissing parameters.ThousandsSeparator,
            TrailingMinusNumbers = boxOrMissing parameters.TrailingMinusNumbers,
            Local = boxOrMissing parameters.Local)

module Worksheet =
    
    /// Adds a worksheet to a workbook and sets its name if specified.
    let addWorksheet (workbook : Workbook) name =
        let worksheet = workbook.Worksheets.Add() :?> Worksheet
        match name with
            | None -> worksheet
            | Some x ->
                worksheet.Name <- x
                worksheet

    /// Adds a worksheet to a workbook before another and sets its name if specified.
    let addWorksheetBefore (workbook : Workbook) worksheet name =
        let worksheet = workbook.Worksheets.Add(Before = worksheet) :?> Worksheet
        match name with
            | None -> worksheet
            | Some x ->
                worksheet.Name <- x
                worksheet

    /// Adds a worksheet to a workbook after another and sets its name if specified.
    let addWorksheetAfter (workbook : Workbook) worksheet name =
        let worksheet = workbook.Worksheets.Add(After = worksheet) :?> Worksheet
        match name with
            | None -> worksheet
            | Some x ->
                worksheet.Name <- x
                worksheet

    /// Adds n worksheets to a workbook.
    let addWorksheets (workbook : Workbook) count =
        workbook.Worksheets.Add(Count = count) :?> Worksheet

    /// Adds n worksheets to a workbook before another.
    let addWorksheetsBefore (workbook : Workbook) worksheet count =
        workbook.Worksheets.Add(Before = worksheet, Count = count) :?> Worksheet

    /// Adds n worksheets to a workbook after another.
    let addWorksheetsAfter (workbook : Workbook) worksheet count =
        workbook.Worksheets.Add(After = worksheet, Count = count) :?> Worksheet

    /// Moves a worksheet to a new workbook.
    let moveSheet2NewBook (worksheet : Worksheet) = worksheet.Move()

    /// Moves a worksheet before another.
    let moveSheetBefore (worksheet : Worksheet) worksheet' = worksheet.Move(Before = worksheet')

    /// Moves a worksheet after another.
    let moveSheetAfter (worksheet : Worksheet) worksheet' = worksheet.Move(After = worksheet')

    /// Copies a worksheet to a new workbook.
    let copySheet (worksheet : Worksheet) = worksheet.Copy()

    /// Copies a worksheet before another.
    let copySheetBefore (worksheet : Worksheet) worksheet' = worksheet.Copy(Before = worksheet')

    /// Copies a worksheet after another.
    let copySheetAfter (worksheet : Worksheet) worksheet' = worksheet.Copy(After = worksheet')

    /// Hides a worksheet.
    let hideWorksheet (worksheet : Worksheet) = worksheet.Visible <- XlSheetVisibility.xlSheetHidden    

    /// Displays a hidden a worksheet.
    let displayHiddenWorksheet (worksheet : Worksheet) = worksheet.Visible <- XlSheetVisibility.xlSheetVisible

    /// Returns a worksheet located at the specified index.
    let worksheetAtIndex (workbook : Workbook) (idx : int) = workbook.Worksheets.[idx] :?> Worksheet

    /// Returns a worksheet having the specified name.    
    let worksheetByName (workbook : Workbook) (name : string) = workbook.Worksheets.[name] :?> Worksheet
    
    /// Activates a worksheet.
    let activateWorksheet (worksheet : Worksheet) = worksheet.Activate()

    /// Activates a worksheet located at the specified index.
    let activateWorksheetByIndex (workbook : Workbook) idx =
        let worksheet = worksheetAtIndex workbook idx
        activateWorksheet worksheet

    /// Activates a worksheet having the specified name.
    let activateWorksheetByName (workbook : Workbook) name =
        let worksheet = worksheetByName workbook name
        activateWorksheet worksheet

    /// Insert a column using the optional shift direction and copy origin parameters.
    let insertColumn (range : Range) shift copy =
        let shift' = boxOrMissing shift
        let copy' = boxOrMissing copy
        range.EntireColumn.Insert(shift', copy') |> ignore

    /// Hides a column.
    let hideColumn (range : Range) = range.EntireColumn.Hidden <- true

    /// Displays a hidden column.
    let displayHiddenColumn (range : Range) = range.EntireColumn.Hidden <- false
    
    /// Deletes a range.
    let deleteCells (range : Range) shift =
        let shift' = boxOrMissing shift
        range.Delete shift' |> ignore

    /// Inserts a range of cells.
    let insertCells (range : Range) shift copy =
        let shift' = boxOrMissing shift
        let copy' = boxOrMissing copy
        range.Insert(shift', copy') |> ignore

    /// Performs an autofill from a source range to a destination one. The two ranges must overlap.
    let autoFillRange (range : Range) range' autoFillType = range.AutoFill(range', autoFillType) |> ignore

module Range =

    /// Selects a range of cells.
    let select (range : Range)= range.Select() |> ignore

    /// Returns a worksheet range.
    let range (worksheet : Worksheet) cell cell' =
        let cell'' = boxOrMissing cell'
        worksheet.Range(cell, cell'')

    /// Copies a range to the clipboard.
    let copyToClipboard (range : Range) = range.Copy() |> ignore

    /// Copies a range to the spedified destination.
    let copyToRange (sourceRange : Range) destinationRange = sourceRange.Copy destinationRange |> ignore
    
    /// Cuts a range.
    let cut (range: Range) = range.Cut() |> ignore

    /// Cuts a range and pastes it into the specified destination.
    let cutPaste (sourceRange: Range) destinationRange = sourceRange.Cut(destinationRange) |> ignore

    /// Returns the range of the column with the specified index.
    let columnAtIndex (worksheet : Worksheet) (idx : int) = worksheet.Columns.[idx] :?> Range |> fun x -> x.EntireColumn

    /// Returns the range of the column with the specified header.    
    let columnByHeader (worksheet : Worksheet) (header : string) = worksheet.Columns.[header] :?> Range |> fun x -> x.EntireColumn

