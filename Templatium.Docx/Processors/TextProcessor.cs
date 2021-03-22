using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Templatium.Docx.Contents;
using Templatium.Docx.Extensions;

namespace Templatium.Docx.Processors
{
    public class TextProcessor : IProcessor
    {
        public Type TargetType => typeof(Content<string>);

        public bool CanFill(
            IDocxTemplater templater, 
            WordprocessingDocument document, 
            SdtElement sdtElement,
            IContent content)
            => content is Content<string>;

        public void FillContent(
            IDocxTemplater templater, 
            WordprocessingDocument document,
            SdtElement sdtElement,
            IContent content)
        {
            var textContent = (Content<string>)content;
            
            var contentBlock = sdtElement.GetFirstOrDefaultByName("sdtContent");
            var textElement = contentBlock.Descendants<Text>().FirstOrDefault();
            if (textElement != null)
                textElement.Text = textContent.Value;
            else
            {
                contentBlock.RemoveAllChildren();
                contentBlock.AppendChild(
                    new Paragraph(
                        new Run(
                            new Text(textContent.Value)
                        )
                    )
                );
            }
        }
    }
}