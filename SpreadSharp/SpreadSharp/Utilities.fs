namespace SpreadSharp

open Microsoft.FSharp.Reflection
open Microsoft.Office.Interop.Excel
open System
open Com

module private Utilities =
    
    let boxOrMissing<'T> = function Some (x : 'T) -> box x | None -> Type.Missing

    let setWorksheetName nameOption (worksheet : Worksheet) =
        match nameOption with
            | None      -> worksheet
            | Some name ->
                worksheet.Name <- name
                worksheet

    let recordFieldsNames recordType =
        FSharpType.GetRecordFields recordType
        |> Array.map (fun x -> box x.Name)

    let fieldsArray records =
        records
        |> Seq.map FSharpValue.GetRecordFields