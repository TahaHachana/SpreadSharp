namespace SpreadSharp

open Microsoft.Office.Interop.Excel
open System.Runtime.InteropServices

module XlApp =

    /// <summary>Returns the stack collection used to hold the Com objects created by the application.
    /// This is the mechanism used to implement proper COM cleanup.</summary>
    /// <returns>The stack collection.</returns>
    let comStack = Com.comStack

    /// <summary>Starts Excel in visible mode.</summary>
    /// <returns>The Excel ApplicationClass instance.</returns>
    let start () =
        ApplicationClass(Visible = true)
        |> Com.pushComObj

    /// <summary>Starts Excel in hidden mode.</summary>
    /// <returns>The created Excel ApplicationClass instance.</returns>
    let startHidden () =
        ApplicationClass(Visible = false)
        |> Com.pushComObj

    
    /// <summary>Returns a reference to an already running Excel instance.</summary>
    /// <returns>The running Excel ApplicationClass instance.</returns>
    let getActiveApp () =
        Marshal.GetActiveObject "Excel.Application"
        :?> Application
        |> Com.pushComObj
    
    /// <summary>Closes Excel and releases its related COM objects.</summary>
    /// <param name="appClass">The Excel ApplicationClass.</param>
    let quit (appClass : ApplicationClass) =
        appClass.Quit ()
        Com.releaseComObjects ()

    /// <summary>Sets the visible property of Excel to false.</summary>
    /// <param name="appClass">The Excel application class instance.</param>
    let hide (appClass : ApplicationClass) =
        appClass.Visible <- false

    /// <summary>Sets the visible property of Excel to true.</summary>
    /// <param name="appClass">The Excel application class instance.</param>
    let unhide (appClass : ApplicationClass) =
        appClass.Visible <- true

    /// <summary>Restores the control of Excel to the user.</summary>
    /// <param name="appClass">The Excel application class instance.</param>
    let restoreUserControl appClass =
        unhide appClass
        appClass.UserControl <- true
        Com.releaseComObjects ()