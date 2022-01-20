namespace Templatium.Docx

open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Wordprocessing
open Templatium.Docx

module DocxTemplater =
    let inline private getAllSdtNodesFromNode (node: OpenXmlElement) =
        OpenXmlHelpers.findDescendantsByName<SdtElement> node Constants.sdt

    let private getAllSdtNodesFromDoc (doc: WordprocessingDocument) =
        let sdts = ResizeArray()

        sdts.AddRange(getAllSdtNodesFromNode doc.MainDocumentPart.Document.Body)

        let inline findSdts (part: OpenXmlPart) =
            OpenXmlHelpers.findDescendantsByName part.RootElement Constants.sdt

        sdts.AddRange(
            doc.MainDocumentPart.HeaderParts
            |> Seq.collect findSdts
        )

        sdts.AddRange(
            doc.MainDocumentPart.FooterParts
            |> Seq.collect findSdts
        )

        sdts.AddRange(findSdts doc.MainDocumentPart.FootnotesPart)
        sdts.AddRange(findSdts doc.MainDocumentPart.EndnotesPart)

        sdts

    let fillNode
        (processors: IProcessor seq)
        (contents: IContent seq)
        (doc: WordprocessingDocument)
        (node: OpenXmlElement)
        =
        let sdts = getAllSdtNodesFromNode node |> Seq.rev

        let metadata =
            { Processors = processors
              Document = doc }

        for sdt in sdts do
            let titleNode =
                OpenXmlHelpers.findFirstNodeByName<SdtAlias> sdt Constants.alias

            match titleNode with
            | None -> ()
            | Some alias ->
                let contentOpt =
                    contents
                    |> Seq.tryFind (fun c -> c.Title = alias.Val.Value)

                match contentOpt with
                | None -> ()
                | Some content ->
                    processors
                    |> Seq.tryFind (fun p -> p.CanFill content sdt metadata)
                    |> Option.iter (fun p -> p.Fill content sdt metadata)

        ()


    let fillDocument (processors: IProcessor seq) (contents: IContent seq) (doc: WordprocessingDocument) =
        fillNode processors contents doc doc.MainDocumentPart.Document.Body

        let fillPart (part: OpenXmlPart) =
            fillNode processors contents doc part.RootElement

        doc.MainDocumentPart.HeaderParts
        |> Seq.iter fillPart

        doc.MainDocumentPart.FooterParts
        |> Seq.iter fillPart

        fillPart doc.MainDocumentPart.FootnotesPart
        fillPart doc.MainDocumentPart.EndnotesPart
        ()

    // TODO: Try find more functional approach
    let deleteContentControls (doc: WordprocessingDocument) =
        let sdts = getAllSdtNodesFromDoc doc

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
