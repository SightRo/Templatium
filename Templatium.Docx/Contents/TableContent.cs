using System.Collections.Generic;
using System.Linq;

namespace Templatium.Docx.Contents
{
    public class TableContent : IContent
    {
        public TableContent(string tagName)
        {
            TagName = tagName;
        }

        public TableContent(string tagName, List<Dictionary<string, IContent>> value)
        {
            TagName = tagName;
            Value = value;
        }

        public TableContent(string tagName, List<List<IContent>> value)
        {
            TagName = tagName;
            Value = value.Select(x => x.ToDictionary(el => el.TagName)).ToList();
        }

        public void AddRow(List<IContent> data)
            => AddRowCore(data);

        public void AddRow(params IContent[] data)
            => AddRowCore(data);

        private void AddRowCore(IEnumerable<IContent> data)
            => Value.Add(data.ToDictionary(d => d.TagName));

        public string TagName { get; init; }
        public List<Dictionary<string, IContent>> Value { get; init; } = new();
        object IContent.Value => Value;
    }
}