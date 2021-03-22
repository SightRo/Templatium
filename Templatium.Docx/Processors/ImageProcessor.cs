using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Templatium.Docx.Contents;
using Templatium.Docx.Extensions;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;

namespace Templatium.Docx.Processors
{
    public class ImageProcessor : IProcessor
    {
        private const int EmuCoefficient = 9525;
        public Type TargetType => typeof(ImageContent);

        public bool CanFill(
            IDocxTemplater templater,
            WordprocessingDocument document,
            SdtElement sdtElement,
            IContent content)
            => content is ImageContent;

        public void FillContent(
            IDocxTemplater templater,
            WordprocessingDocument document,
            SdtElement sdtElement,
            IContent content)
        {
            var imageContent = (ImageContent) content;

            var contentBlock = sdtElement.GetFirstOrDefaultByName("sdtContent");
            var imageBlock = contentBlock.Descendants<Drawing>().FirstOrDefault();

            if (imageBlock != null)
                ReplaceImage(document, contentBlock, imageContent);
            else
                InsertImage(document, sdtElement, imageContent);
        }

        private void ReplaceImage(WordprocessingDocument document, OpenXmlElement contentBlock, ImageContent content)
        {
            var imageId = contentBlock.Descendants<Blip>().First().Embed.Value;
            var curImage = (ImagePart) document.MainDocumentPart.GetPartById(imageId);

            if (content.Height != null && content.Width != null)
            {
                var extents = contentBlock.Descendants<DW.Extent>();
                foreach (var extent in extents)
                {
                    extent.Cx = content.Width * EmuCoefficient;
                    extent.Cy = content.Height * EmuCoefficient;
                }
            }

            using var writer = new BinaryWriter(curImage.GetStream());
            writer.Write(content.Value);
        }

        private void InsertImage(WordprocessingDocument document, SdtElement element, ImageContent content)
        {
            var type = GetImageType(content.Extension);
            if (type == null)
                throw new NotSupportedException("Specified image extension is not supported");

            if (content.Height == null || content.Width == null)
                throw new ArgumentException("Image sizes must be specified when sdt element have no image example");

            var imagePart = document.MainDocumentPart.AddImagePart(type.Value);

            using (var ms = new MemoryStream(content.Value))
            {
                imagePart.FeedData(ms);
            }

            var relationId = document.MainDocumentPart.GetIdOfPart(imagePart);

            Int64Value width = content.Width * EmuCoefficient;
            Int64Value height = content.Height * EmuCoefficient;

            var draw = new Drawing(
                new DW.Inline(
                    new DW.Extent {Cx = width, Cy = height},
                    new DW.EffectExtent
                    {
                        LeftEdge = 0L,
                        TopEdge = 0L,
                        RightEdge = 0L,
                        BottomEdge = 0L
                    },
                    new DW.DocProperties
                    {
                        Id = (UInt32Value) 1U,
                        Name = $"picture{DateTime.Now.Ticks}.{content.Extension}"
                    },
                    new DW.NonVisualGraphicFrameDrawingProperties(new GraphicFrameLocks {NoChangeAspect = true}),
                    new Graphic(
                        new GraphicData(
                                new PIC.Picture(
                                    new PIC.NonVisualPictureProperties(
                                        new PIC.NonVisualDrawingProperties
                                        {
                                            Id = (UInt32Value) 0U,
                                            Name = relationId
                                        },
                                        new PIC.NonVisualPictureDrawingProperties()),
                                    new PIC.BlipFill(
                                        new Blip(
                                            new BlipExtensionList(
                                                new BlipExtension {Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"})
                                        )
                                        {
                                            Embed = relationId,
                                            CompressionState =
                                                BlipCompressionValues.Print
                                        },
                                        new Stretch(
                                            new FillRectangle())),
                                    new PIC.ShapeProperties(
                                        new Transform2D(
                                            new Offset {X = 0L, Y = 0L},
                                            new Extents {Cx = width, Cy = height}),
                                        new PresetGeometry(new AdjustValueList())
                                            {Preset = ShapeTypeValues.Rectangle})))
                            {Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"})
                )
                {
                    DistanceFromTop = (UInt32Value) 0U,
                    DistanceFromBottom = (UInt32Value) 0U,
                    DistanceFromLeft = (UInt32Value) 0U,
                    DistanceFromRight = (UInt32Value) 0U,
                    EditId = "50D07946"
                });
            var contentBlock = element.Descendants<SdtContentBlock>().FirstOrDefault();
            contentBlock.RemoveAllChildren();
            contentBlock.AppendChild(new Paragraph(new Run(draw)));
            //element.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>().First().AppendChild(draw);
        }

        private ImagePartType? GetImageType(string extension)
            => extension.ToLower() switch
            {
                ".bmp" => ImagePartType.Bmp,
                ".emf" => ImagePartType.Emf,
                ".gif" => ImagePartType.Gif,
                ".ico" => ImagePartType.Icon,
                ".icon" => ImagePartType.Icon,
                ".jpg" => ImagePartType.Jpeg,
                ".jpeg" => ImagePartType.Jpeg,
                ".pcx" => ImagePartType.Pcx,
                ".png" => ImagePartType.Png,
                ".tif" => ImagePartType.Tiff,
                ".tiff" => ImagePartType.Tiff,
                ".wmf" => ImagePartType.Wmf,
                _ => null
            };
    }
}