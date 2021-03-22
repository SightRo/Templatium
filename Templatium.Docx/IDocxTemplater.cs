using System;
using DocumentFormat.OpenXml.Packaging;
using Templatium.Docx.Contents;
using Templatium.Docx.Processors;

namespace Templatium.Docx
{
    public interface IDocxTemplater
    {
        IProcessor? GetProcessor(Type targetType);
        void AddProcessor(IProcessor processor);
        void GenerateDocument(WordprocessingDocument input, ContentContainer container);
        void DeleteContentControls(WordprocessingDocument doc);
    }
}