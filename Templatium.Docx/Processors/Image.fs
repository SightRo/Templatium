namespace Templatium.Docx.Processors

open Templatium.Docx
open DocumentFormat.OpenXml.Wordprocessing
open System.Linq
open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Drawing
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

type ImageSize =
    | Original
    | Size of width: int64 * height: int64

type ImageContent =
    { Title: string
      Image: byte array
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
    let getUniqId32 () = 12u
    let getUniqId64 () = 12UL

    let replaceImage (doc: WordprocessingDocument) (imageBlock: OpenXmlElement) content =
        let imageId =
            imageBlock.Descendants<Blip>().First().Embed.Value

        let imagePart = doc.MainDocumentPart.GetPartById imageId

        match content.Size with
        | Size (width, height) ->
            let extent =
                imageBlock.Descendants<Wordprocessing.Extent>().FirstOrDefault()

            extent.Cx <- Int64Value(width)
            extent.Cy <- Int64Value(height)
        | Original -> ()

        use writer = new BinaryWriter(imagePart.GetStream())
        writer.Write(content.Image)
        ()

    let insertImage (doc: WordprocessingDocument) (contentBlock: OpenXmlElement) content =
        let imageType = convertToWordType content.Type

        match content.Size with
        | Original -> failwith "lol"
        | Size (width, height) ->
            let imagePart =
                doc.MainDocumentPart.AddImagePart imageType

            use ms = new MemoryStream(content.Image)
            imagePart.FeedData ms

            let relationId =
                doc.MainDocumentPart.GetIdOfPart imagePart

            // Good luck debugging this shit
            // https://docs.microsoft.com/en-us/office/open-xml/how-to-insert-a-picture-into-a-word-processing-document
            let draw =
                Drawing(
                    Wordprocessing.Inline(
                        Wordprocessing.Extent(Cx = Int64Value(width), Cy = Int64Value(height)),
                        Wordprocessing.EffectExtent(
                            LeftEdge = Int64Value(0),
                            TopEdge = Int64Value(0),
                            RightEdge = Int64Value(0),
                            BottomEdge = Int64Value(0)
                        ),
                        Wordprocessing.DocProperties(
                            Id = UInt32Value(getUniqId32 ()),
                            Name = StringValue(DateTime.Now.Ticks.ToString())
                        ),
                        Wordprocessing.NonVisualGraphicFrameDrawingProperties(GraphicFrameLocks(NoChangeAspect = true)),
                        Graphic(
                            GraphicData(
                                Pictures.Picture(
                                    Pictures.NonVisualPictureProperties(
                                        Pictures.NonVisualDrawingProperties(
                                            Id = UInt32Value(getUniqId32 ()),
                                            Name = relationId
                                        ),
                                        Pictures.NonVisualPictureDrawingProperties()
                                    ),
                                    Pictures.BlipFill(
                                        Blip(
                                            BlipExtensionList(
                                                BlipExtension(
                                                    Uri = StringValue("{28A0092B-C50C-407E-A947-70E740481C1C}")
                                                )
                                            ),
                                            Embed = StringValue(relationId),
                                            CompressionState = BlipCompressionValues.Print
                                        ),
                                        Stretch(FillRectangle())
                                    ),
                                    Pictures.ShapeProperties(
                                        Transform2D(
                                            Offset(X = Int64Value(0), Y = Int64Value(0)),
                                            Extents(Cx = Int64Value(width), Cy = Int64Value(height))
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

        ()

type ImageProcessor() =
    interface IProcessor with
        member _.CanFill content _ _  = content :? ImageContent

        member _.Fill content sdt metadata =
            let imageContent = content :?> ImageContent

            let contentBlock =
                OpenXmlHelpers.findFirstNodeByName sdt Constants.sdtContent

            match contentBlock with
            | Some block ->
                let imageBlock =
                    block.Descendants<Drawing>().FirstOrDefault()

                match imageBlock with
                | null -> ImageProcessor.insertImage metadata.Document block imageContent
                | _ -> ImageProcessor.replaceImage metadata.Document imageBlock imageContent
            | _ -> ()

            ()
