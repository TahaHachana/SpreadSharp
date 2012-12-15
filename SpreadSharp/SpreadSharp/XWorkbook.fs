namespace SpreadSharp

open Microsoft.Office.Interop.Excel

module XWorkbook =

    /// <summary>Adds a workbook to an Excel app.</summary>
    /// <returns>The new workbook.</returns>
    let add (appClass : ApplicationClass) =
        appClass.Workbooks.Add()
        |> COM.pushComObj

    /// <summary>Closes a workbook.</summary>
    /// <param name="workbook">The workbook to close.</param>
    /// <param name="saveChanges">The SaveChanges setting.</param>
    /// <param name="fileName">The saving path.</param>
    /// <param name="routeWorkbook">The RouteWorkbook setting.</param>
    /// <returns>unit</returns>
    let close (workbook : Workbook) saveChanges fileName routeWorkbook =
        workbook.Close(
            SaveChanges   = Utilities.boxOrMissing<bool>   saveChanges,
            Filename      = Utilities.boxOrMissing<string> fileName,
            RouteWorkbook = Utilities.boxOrMissing<bool>   routeWorkbook
        )

    /// <summary>Opens an existing workbook.</summary>
    /// <param name="appClass">The Excel ApplicationClass.</param>
    /// <param name="fileName">The name of the workbook file.</param>
    /// <returns>The opened workbook.</returns>
    let openWorkbook (appClass : ApplicationClass) fileName =
        appClass.Workbooks.Open fileName
        |> COM.pushComObj

    /// <summary>Saves a workbook in the MyDocuments folder.</summary>
    /// <param name="workbook">The workbook to save.</param>
    /// <returns>unit</returns>
    let save (workbook : Workbook) = workbook.Save()

    /// <summary>Saves a workbook using the specified file name.</summary>
    /// <param name="workbook">The workbook to save.</param>
    /// <param name="fileName">The name of the workbook file.</param>
    /// <returns>unit</returns>
    let saveAs (workbook : Workbook) (fileName : string) = workbook.SaveAs(Filename = fileName)

//    type OpenTextParameters =
//        {
//            Filename             : string
//            Origin               : TextFileOrigin option
//            StartRow             : int option
//            DataType             : XlTextParsingType option
//            TextQualifier        : XlTextQualifier
//            ConsecutiveDelimiter : bool option
//            Tab                  : bool option
//            Semicolon            : bool option
//            Comma                : bool option
//            Space                : bool option
//            Other                : bool option
//            OtherChar            : bool option
//            FieldInfo            : int [,] option
//            TextVisualLayout     : obj option
//            DecimalSeparator     : obj option
//            ThousandsSeparator   : obj option
//            TrailingMinusNumbers : bool option
//            Local                : bool option
//        }

//    /// Opens a text file as a new workbook with a single sheet.
//    let openText (application : ApplicationClass) (parameters : OpenTextParameters) =
//        application.Workbooks.OpenText(
//            Filename = parameters.Filename,
//            Origin = (parameters.Origin |> function
//                | Some x -> x |> function Platform p -> box p | CodePage cp -> box cp
//                | None -> Type.Missing),
//            StartRow = boxOrMissing parameters.StartRow,
//            DataType = boxOrMissing parameters.DataType,
//            TextQualifier = parameters.TextQualifier,
//            ConsecutiveDelimiter = boxOrMissing parameters.ConsecutiveDelimiter,
//            Tab = boxOrMissing parameters.Tab,
//            Semicolon = boxOrMissing parameters.Semicolon,
//            Comma = boxOrMissing parameters.Comma,
//            Space = boxOrMissing parameters.Space,
//            Other = boxOrMissing parameters.Other,
//            OtherChar = parameters.OtherChar,
//            FieldInfo = boxOrMissing parameters.FieldInfo,
//            TextVisualLayout = boxOrMissing parameters.TextVisualLayout,
//            DecimalSeparator = boxOrMissing parameters.DecimalSeparator,
//            ThousandsSeparator = boxOrMissing parameters.ThousandsSeparator,
//            TrailingMinusNumbers = boxOrMissing parameters.TrailingMinusNumbers,
//            Local = boxOrMissing parameters.Local)