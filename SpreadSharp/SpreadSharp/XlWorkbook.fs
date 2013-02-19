namespace SpreadSharp

open Microsoft.Office.Interop.Excel

module XlWorkbook =

    /// <summary>Adds a workbook to an Excel app.</summary>
    /// <returns>The new workbook.</returns>
    let add (appClass : ApplicationClass) =
        appClass.Workbooks.Add()
        |> Com.pushComObj

    /// <summary>Closes a workbook. Use the save and saveAs function to save
    /// a workbook before closing it.</summary>
    /// <param name="workbook">The workbook to close.</param>
    let close (workbook : Workbook) = workbook.Close()

    /// <summary>Opens an existing workbook.</summary>
    /// <param name="appClass">The Excel ApplicationClass.</param>
    /// <param name="fileName">The name of the workbook file.</param>
    /// <returns>The opened workbook.</returns>
    let openWorkbook (appClass : ApplicationClass) fileName =
        appClass.Workbooks.Open fileName
        |> Com.pushComObj

    /// <summary>Saves a workbook in the MyDocuments folder.</summary>
    /// <param name="workbook">The workbook to save.</param>
    let save (workbook : Workbook) = workbook.Save()

    /// <summary>Saves a workbook using the specified file name.</summary>
    /// <param name="workbook">The workbook to save.</param>
    /// <param name="fileName">The name of the workbook file.</param>
    let saveAs (workbook : Workbook) (fileName : string) = workbook.SaveAs(Filename = fileName)