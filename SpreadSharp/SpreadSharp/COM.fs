namespace SpreadSharp

open System
open System.Collections.Generic
open System.Runtime.InteropServices

module Com =

    /// <summary>Holds COM objects explicitly created by the application.
    /// This is the mechanism used to hide proper COM cleanup.</summary>
    let comStack = Stack<obj>()

    /// <summary>Inserts a COM object created without calling a library function
    /// in the stack collection dedicated to holding them. Adding the object ensures
    /// that it's properly released when XApp.quit is executed.</summary>
    /// <param name="comObj">The COM object created without using the modules of the library.</param>
    /// <returns>The object itself.</returns>
    let pushComObj comObj =
        comStack.Push comObj
        comObj

    let inline private finalRelease comObj = Marshal.FinalReleaseComObject comObj |> ignore

    let inline private collectGarbage () =
        GC.Collect ()
        GC.WaitForPendingFinalizers ()
    
    /// <summary>Releases COM objects explicitly created by the application.</summary>
    let releaseComObjects () =  
        comStack |> Seq.iter finalRelease
        comStack.Clear ()
        collectGarbage ()