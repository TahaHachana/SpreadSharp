namespace SpreadSharp

open Microsoft.Office.Interop.Excel
open System
open System.Runtime.InteropServices
open COM

module private Utilities =
    
    let releaseComObjects () =  
        comStack |> Seq.iter (fun x -> Marshal.FinalReleaseComObject x |> ignore)
        comStack.Clear()

    let collectGarbage () =
        GC.Collect ()
        GC.WaitForPendingFinalizers ()

    let boxOrMissing<'T> = function Some (x : 'T) -> box x | None -> Type.Missing

    let setWorksheetName nameOption (worksheet : Worksheet) =
        match nameOption with
            | None      -> worksheet
            | Some name ->
                worksheet.Name <- name
                worksheet