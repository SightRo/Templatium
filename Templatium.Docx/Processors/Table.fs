namespace Templatium.Docx.Processors

open DocumentFormat.OpenXml
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


type TableProcessor() =
    interface IProcessor with
        member _.CanFill content _ _ = content :? TableContent

        member _.Fill content sdt metadata =
            let tableContent = content :?> TableContent

            OpenXmlHelpers.findFirstNodeByName sdt Constants.sdtContent
            |> Option.bind (fun contentBlock -> OpenXmlHelpers.findFirstNodeByName<Table> contentBlock Constants.table)
            |> Option.bind
                (fun tableNode ->
                    tableNode.Descendants<TableRow>()
                    |> Seq.tryFind
                        (fun r ->
                            match OpenXmlHelpers.findFirstNodeByName r Constants.sdt with
                            | Some _ -> true
                            | None -> false))
            |> Option.iter
                (fun rowTemplate ->
                    let mutable previousRow = rowTemplate :> OpenXmlElement

                    for contentRow in tableContent.Rows do
                        let clonedRow = rowTemplate.CloneNode true

                        DocxTemplater.fillNode metadata.Processors contentRow metadata.Document clonedRow
                        previousRow <- previousRow.InsertAfterSelf clonedRow

                    rowTemplate.Remove())

            ()
