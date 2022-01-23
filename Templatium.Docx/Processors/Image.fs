namespace Templatium.Docx.Processors

open DocumentFormat.OpenXml.Drawing
open DocumentFormat.OpenXml.Drawing.Wordprocessing
open Templatium.Docx
open DocumentFormat.OpenXml.Wordprocessing
open System.Linq
open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open System.IO
open System

type ImageType =
    | Png
    | Jpeg
    | Gif
    | Bmp
    | Emf
    | Icon
    | Pcx
    | Tiff
    | Wmf

type ImageSize = { Width: int64; Height: int64 }

type ImageFormat =
    | Original
    | Size of ImageSize

type ImagePartBehavior =
    | Add
    | Replace

type ImageReplaceContent =
    { Title: string
      Image: Stream
      Type: ImageType
      Format: ImageFormat
      ImagePartBehavior: ImagePartBehavior }
    interface IContent with
        member this.Title = this.Title
        member this.Value = this.Image

type ImageAddContent =
    { Title: string
      Image: Stream
      Type: ImageType
      Size: ImageSize }
    interface IContent with
        member this.Title = this.Title
        member this.Value = this.Image

module internal ImageProcessor =
    let convertToWordType =
        function
        | Png -> ImagePartType.Png
        | Jpeg -> ImagePartType.Jpeg
        | Gif -> ImagePartType.Gif
        | Bmp -> ImagePartType.Bmp
        | Emf -> ImagePartType.Emf
        | Icon -> ImagePartType.Icon
        | Pcx -> ImagePartType.Pcx
        | Tiff -> ImagePartType.Tiff
        | Wmf -> ImagePartType.Wmf

    // TODO: Investigate. Does it really need id properties
    let getUniqId32 () = Convert.ToUInt32(Random.Shared.Next())

    let getUniqId64 () =
        Convert.ToUInt64(Random.Shared.NextInt64())

    let addImagePart (doc: WordprocessingDocument) (imageType: ImagePartType) image =
        let imagePart =
            doc.MainDocumentPart.AddImagePart imageType

        imagePart.FeedData image
        imagePart

    let replaceImagePart (doc: WordprocessingDocument) imageId image =
        let imagePart = doc.MainDocumentPart.GetPartById imageId
        imagePart.FeedData image

    let replaceImage (doc: WordprocessingDocument) (imageBlock: OpenXmlElement) (content: ImageReplaceContent) =
        let imageType = convertToWordType content.Type
        let blip = imageBlock.Descendants<Blip>().First()

        match content.ImagePartBehavior with
        | Add ->
            let imagePart = addImagePart doc imageType content.Image
            blip.Embed <- doc.MainDocumentPart.GetIdOfPart imagePart
        | Replace ->
            let imageId = blip.Embed.Value
            replaceImagePart doc imageId content.Image


        match content.Format with
        | Size size ->
            imageBlock.Descendants<Extent>()
            |> Seq.iter
                (fun extent ->
                    extent.Cx <- Int64Value(size.Width)
                    extent.Cy <- Int64Value(size.Height))
        | Original -> ()

        ()

    let insertImage (doc: WordprocessingDocument) (contentBlock: OpenXmlElement) (content: ImageAddContent) =
        let imageType = convertToWordType content.Type

        let imagePart =
            doc.MainDocumentPart.AddImagePart imageType

        imagePart.FeedData content.Image

        let relationId =
            doc.MainDocumentPart.GetIdOfPart imagePart

        // Good luck debugging this shit
        // https://docs.microsoft.com/en-us/office/open-xml/how-to-insert-a-picture-into-a-word-processing-document
        let draw =
            Drawing(
                Inline(
                    Extent(Cx = Int64Value(content.Size.Width), Cy = Int64Value(content.Size.Height)),
                    EffectExtent(
                        LeftEdge = Int64Value(0),
                        TopEdge = Int64Value(0),
                        RightEdge = Int64Value(0),
                        BottomEdge = Int64Value(0)
                    ),
                    DocProperties(Id = UInt32Value(getUniqId32 ()), Name = StringValue(DateTime.Now.Ticks.ToString())),
                    NonVisualGraphicFrameDrawingProperties(GraphicFrameLocks(NoChangeAspect = true)),
                    Graphic(
                        GraphicData(
                            Drawing.Pictures.Picture(
                                Drawing.Pictures.NonVisualPictureProperties(
                                    Drawing.Pictures.NonVisualDrawingProperties(
                                        Id = UInt32Value(getUniqId32 ()),
                                        Name = relationId
                                    ),
                                    Drawing.Pictures.NonVisualPictureDrawingProperties()
                                ),
                                Drawing.Pictures.BlipFill(
                                    Blip(
                                        BlipExtensionList(
                                            BlipExtension(Uri = StringValue("{28A0092B-C50C-407E-A947-70E740481C1C}"))
                                        ),
                                        Embed = StringValue(relationId),
                                        CompressionState = BlipCompressionValues.Print
                                    ),
                                    Stretch(FillRectangle())
                                ),
                                Drawing.Pictures.ShapeProperties(
                                    Transform2D(
                                        Offset(X = Int64Value(0), Y = Int64Value(0)),
                                        Extents(
                                            Cx = Int64Value(content.Size.Width),
                                            Cy = Int64Value(content.Size.Height)
                                        )
                                    ),
                                    PresetGeometry(AdjustValueList(), Preset = ShapeTypeValues.Rectangle)
                                )
                            ),
                            Uri = StringValue("http://schemas.openxmlformats.org/drawingml/2006/picture")
                        )
                    ),
                    DistanceFromTop = UInt32Value(0u),
                    DistanceFromBottom = UInt32Value(0u),
                    DistanceFromLeft = UInt32Value(0u),
                    DistanceFromRight = UInt32Value(0u),
                    EditId = "50D07946"
                )
            )

        contentBlock.RemoveAllChildren()

        contentBlock.AppendChild(Paragraph(Run(draw)))
        |> ignore

        ()

type ImageProcessor() =
    interface IProcessor with
        member _.CanFill content _ _ =
            match content with
            | :? ImageReplaceContent -> true
            | :? ImageAddContent -> true
            | _ -> false

        member _.Fill content sdt metadata =
            let contentBlock =
                OpenXmlHelpers.findFirstNodeByName sdt Constants.sdtContent

            match contentBlock with
            | Some block ->
                let imageBlock =
                    block.Descendants<Drawing>().FirstOrDefault()

                match imageBlock, content with
                | null, (:? ImageAddContent as addContent) ->
                    ImageProcessor.insertImage metadata.Document block addContent
                | _node, (:? ImageReplaceContent as repContent) ->
                    ImageProcessor.replaceImage metadata.Document imageBlock repContent
                | _ -> ()
            | _ -> ()

            ()
