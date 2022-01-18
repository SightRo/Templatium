namespace Templatium.Docx

open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Wordprocessing
open System.Linq
open System.Collections.Generic
open Templatium.Docx

module DocxTemplater =

    let private getAllSdtNodesFromNode (node: OpenXmlElement) =
        let sdts =
            OpenXmlHelpers.findAllNodeByName node Constants.sdt

        sdts.OfType<SdtElement>()

    // TODO: Handle headers and footers
    let private getAllSdtNodesFromDoc (doc: WordprocessingDocument) =
        getAllSdtNodesFromNode doc.MainDocumentPart.Document.Body

    let fillNode
        (processors: IProcessor seq)
        (contents: IContent seq)
        (doc: WordprocessingDocument)
        (node: OpenXmlElement)
        =
        let sdts = getAllSdtNodesFromNode node |> Seq.rev

        for sdt in sdts do
            let titleNode =
                OpenXmlHelpers.findFirstNodeByName<SdtAlias> sdt Constants.alias

            match titleNode with
            | None -> ()
            | Some alias ->
                let contentOpt =
                    contents
                    |> Seq.tryFind (fun c -> c.Title = alias.Val)

                match contentOpt with
                | None -> ()
                | Some content ->
                    processors
                    |> Seq.tryFind (fun p -> p.CanFill doc sdt content)
                    |> Option.iter (fun p -> p.Fill doc sdt content)

        ()


    let fillDocument (processors: IProcessor seq) (contents: IContent seq) (doc: WordprocessingDocument) =
        let allSdtNodes = getAllSdtNodesFromDoc doc

        ()

    // TODO: Try find more functional approach
    let deleteContentControls (doc: WordprocessingDocument) =
        let sdts = List(getAllSdtNodesFromDoc doc)

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
