using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Templatium.Docx.Contents;
using Templatium.Docx.Extensions;
using Templatium.Docx.Processors;

namespace Templatium.Docx
{
    public class DocxTemplater : IDocxTemplater
    {
        private readonly Dictionary<Type, IProcessor> _processors = new();

        public IReadOnlyDictionary<Type, IProcessor> Processors => _processors;

        public IProcessor? GetProcessor(Type targetType)
            => _processors.TryGetValue(targetType, out var processor)
                ? processor
                : null;

        public void AddProcessor(IProcessor processor)
            => _processors.Add(processor.TargetType, processor);

        public void GenerateDocument(WordprocessingDocument input, ContentContainer contents)
        {
            var collection = GetAllSdtElements(input);

            foreach (var sdt in collection)
            {
                var tag = sdt.SdtProperties.Descendants<Tag>().FirstOrDefault();
                if (tag == null)
                    throw new Exception("Not found any tag inside std element");

                var content = contents.GetContent(tag.Val.Value);
                if (content == null)
                    continue;

                if (!_processors.TryGetValue(content.GetType(), out var processor))
                    throw new Exception($"Not found processor for {content.GetType()}");

                if (processor.CanFill(this, input, sdt, content))
                    processor.FillContent(this, input, sdt, content);
            }
        }

        public void DeleteContentControls(WordprocessingDocument doc)
        {
            var sdtElements = GetAllSdtElements(doc);

            for (var i = sdtElements.Count - 1; i >= 0; i--)
            {
                var sdt = sdtElements[i];

                var content = sdt.Descendants().FirstOrDefault(o => o.LocalName == "sdtContent");

                if (content != null)
                {
                    var prev = sdt.PreviousSibling();
                    foreach (var child in content.ChildElements)
                        prev = prev.InsertAfterSelf((OpenXmlElement) child.Clone());
                }

                sdt.Remove();
            }
        }

        private List<SdtElement> GetAllSdtElements(WordprocessingDocument doc)
        {
            var collection = doc.MainDocumentPart.Document.Body.FindByName("sdt").ToList();

            collection.AddRange(
                doc.MainDocumentPart.FooterParts
                    .SelectMany(c => c.Footer.Descendants<SdtElement>())
            );
            collection.AddRange(
                doc.MainDocumentPart.HeaderParts
                    .SelectMany(c => c.Header.Descendants<SdtElement>())
            );

            return collection.OfType<SdtElement>().ToList();
        }
    }
}