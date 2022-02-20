namespace Templatium.Docx.Processors

open Templatium.Docx
open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Wordprocessing
open Microsoft.FSharp.Collections
open Templatium.Docx

type ListContent =
    { Title: string
      Items: List<IContent> }
    interface IContent with
        member this.Title = this.Title
        member this.Value = this.Items


type ListProcessor() =
    interface IProcessor with
        member _.CanFill content _ _ = content :? ListContent

        member _.Fill content sdt metadata =
            let listContent = content :?> ListContent

            let contentNodeOpt =
                OpenXmlHelpers.findFirstNodeByName sdt Constants.sdtContent

            match contentNodeOpt with
            | None -> ()
            | Some contentNode ->
                let templateNodeOpt: OpenXmlElement option =
                    OpenXmlHelpers.findFirstNodeByName contentNode Constants.sdt

                match templateNodeOpt with
                | None -> ()
                | Some templateNode ->
                    let mutable previousItemNode = templateNode

                    for itemContent in listContent.Items do
                        let clonedNode = templateNode.CloneNode true
                        previousItemNode <- previousItemNode.InsertAfterSelf clonedNode
                        DocxTemplater.fillNode metadata.Processors [ itemContent ] metadata.Document clonedNode

                    templateNode.Remove()
                    ()

                ()

            ()
