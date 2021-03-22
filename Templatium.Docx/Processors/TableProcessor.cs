using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Templatium.Docx.Contents;
using Templatium.Docx.Extensions;

namespace Templatium.Docx.Processors
{
    public class TableProcessor : IProcessor
    {
        public Type TargetType => typeof(TableContent);

        public bool CanFill(
            IDocxTemplater templater,
            WordprocessingDocument document,
            SdtElement sdtElement,
            IContent content)
            => content is TableContent;

        public void FillContent(
            IDocxTemplater templater,
            WordprocessingDocument document,
            SdtElement sdtElement,
            IContent content)
        {
            var tableContent = (TableContent)content;

            var contentBlock = sdtElement.GetFirstOrDefaultByName("sdtContent");
            var table = contentBlock.Descendants<Table>().FirstOrDefault();
            if (table == null)
                throw new Exception("Not found table");

            var rows = table.Descendants<TableRow>().ToList();
            var startRowCount = rows.Count;
            if (startRowCount == 0)
                throw new Exception("Table have no rows");

            var currentContentIndex = 0;
            TableRow? rowToClone = null;
            
            for (var index = 0; index <= rows.Count; index++)
            {
                if (tableContent.Value.Count <= currentContentIndex)
                    break;

                if (index == rows.Count)
                {
                    if (rowToClone == null)
                        throw new Exception("Incorrect table format");
                    rows.Add((TableRow)rowToClone.CloneNode(true));
                }

                var isSomethingAdded = false;

                foreach (var cell in rows[index].ChildElements<TableCell>())
                {
                    var sdt = cell.Descendants<SdtElement>().FirstOrDefault();
                    if (sdt == null)
                    {
                        if (currentContentIndex != 0)
                            throw new Exception("Incorrect table format");
                        continue;
                    }

                    if (rowToClone == null)
                        rowToClone = rows[index];

                    var tag = sdt.SdtProperties.Descendants<Tag>().First();
                    if (tag == null)
                        throw new Exception("Not found any tag inside std table element");

                    if (!tableContent.Value[currentContentIndex].TryGetValue(tag.Val.Value, out var value))
                        throw new Exception("Not found appropriate value for table cell");

                    var processor = templater.GetProcessor(value.GetType());
                    if (processor == null)
                        throw new Exception($"Not found processor for {value.GetType()}");

                    if (processor.CanFill(templater, document, sdt, value))
                        processor.FillContent(templater, document, sdt, value);

                    isSomethingAdded = true;
                }

                currentContentIndex += isSomethingAdded ? 1 : 0;
            }
            
            for (int i = startRowCount ; i < rows.Count; i++)
            {
                rows[i-1].InsertAfterSelf(rows[i]);
            }
        }
    }
}