using System;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Templatium.Docx.Contents;

namespace Templatium.Docx.Processors
{
    public interface IProcessor
    {
        Type TargetType { get; }

        bool CanFill(
            IDocxTemplater templater,
            WordprocessingDocument document,
            SdtElement sdtElement,
            IContent content);

        void FillContent(
            IDocxTemplater templater,
            WordprocessingDocument document,
            SdtElement sdtElement,
            IContent content);
    }
}