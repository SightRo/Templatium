namespace Templatium.Docx

open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Wordprocessing
open System.Linq
open System.Collections.Generic
open Templatium.Docx

module DocxTemplater =

    // TODO: Handle headers and footers
    let private getAllSdtElements (doc: WordprocessingDocument) =
        let sdts =
            OpenXmlHelpers.findAllNodeByName doc.MainDocumentPart.Document.Body Constants.sdt

        sdts.OfType<SdtElement>()

    // TODO: Doesn't work. Investigate.
    let generateDocument (processors: IProcessor seq) (contents: IContent seq) (doc: WordprocessingDocument) =
        let sdts = getAllSdtElements doc

        for sdt in sdts do
            let titleNode =
                OpenXmlHelpers.findFirstNodeByName<SdtAlias> sdt Constants.alias

            match titleNode with
            | None -> ()
            | Some node ->
                let contentOpt =
                    contents
                    |> Seq.tryFind (fun c -> c.Title = node.Val)

                match contentOpt with
                | None -> ()
                | Some content ->
                    let processorOpt =
                        processors
                        |> Seq.tryFind (fun p -> p.CanFill doc sdt content)

                    match processorOpt with
                    | None -> ()
                    | Some processor -> processor.Fill doc sdt content

        ()

    // TODO: Try find more functional approach
    let deleteContentControls (doc: WordprocessingDocument) =
        let sdts = List(getAllSdtElements doc)

        for i = sdts.Count - 1 downto 0 do
            let sdt = sdts[i]

            let sdtContent =
                OpenXmlHelpers.findFirstNodeByName sdt Constants.sdtContent

            match sdtContent with
            | Some block ->
                let mutable prev = block.PreviousSibling()

                for child in block.ChildElements do
                    prev <- prev.InsertAfterSelf(child.CloneNode(true))
            | None -> ()

            sdt.Remove()

        ()
