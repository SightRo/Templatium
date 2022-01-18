namespace Templatium.Docx.Processors

open System.Collections.Generic
open DocumentFormat.OpenXml.Wordprocessing
open Microsoft.FSharp.Collections
open Templatium.Docx
open System.Linq

type TableContent =
    { Title: string
      Rows: List<IContent> }
    interface IContent with
        member this.Title = this.Title
        member this.Value = this.Rows
        

type TableProcessor =
    interface IProcessor with
        member _.CanFill _ _ content = content :? TableContent

        member _.Fill _ sdt content =
            let tableContent = content :?> TableContent
            OpenXmlHelpers.findFirstNodeByName sdt Constants.sdtContent
            |> Option.bind (fun contentBlock ->
                OpenXmlHelpers.findFirstNodeByName<Table> contentBlock Constants.table)
            |> Option.bind (fun tableNode ->
                let rows = tableNode.Descendants<TableRow>()
                match rows.Count() with
                | 0 -> None
                | 1 -> Some row
                | _ -> None)
            |> Option.bind (fun row ->
                    
                )
            |> ignore
            
            ()
            
