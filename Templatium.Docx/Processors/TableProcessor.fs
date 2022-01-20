namespace Templatium.Docx.Processors

open DocumentFormat.OpenXml.Wordprocessing
open Microsoft.FSharp.Collections
open Templatium.Docx
open System.Linq

type TableContent =
    { Title: string
      Rows: List<List<IContent>> }
    interface IContent with
        member this.Title = this.Title
        member this.Value = this.Rows


type TableProcessor =
    interface IProcessor with
        member _.CanFill content _ _ = content :? TableContent

        member _.Fill content sdt metadata =
            let tableContent = content :?> TableContent

            OpenXmlHelpers.findFirstNodeByName sdt Constants.sdtContent
            |> Option.bind (fun contentBlock -> OpenXmlHelpers.findFirstNodeByName<Table> contentBlock Constants.table)
            |> Option.bind
                (fun tableNode ->
                    let rows = tableNode.Descendants<TableRow>()

                    match rows.Count() with
                    | 0 -> None
                    | 1 -> Some(rows.First())
                    | _ -> None)
            |> Option.iter
                (fun row ->
                    let mutable previousRow = row

                    for contentRow in tableContent.Rows do
                        let clonedRow = (row.CloneNode true) :?> TableRow
                        DocxTemplater.fillNode metadata.Processors contentRow metadata.Document clonedRow
                        previousRow <- previousRow.InsertAfterSelf clonedRow
                    
                    row.Remove())
            ()
