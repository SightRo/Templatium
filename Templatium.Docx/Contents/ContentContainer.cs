using System.Collections.Generic;

namespace Templatium.Docx.Contents
{
    public class ContentContainer
    {
        private readonly Dictionary<string, IContent> _contents = new();
        public IReadOnlyDictionary<string, IContent> Contents => _contents;

        public IContent? GetContent(string tagName)
            => _contents.TryGetValue(tagName, out var content)
                ? content
                : null;
        
        public void AddContent(IContent content)
            => _contents.Add(content.TagName, content);

        public void AddContent<TVal>(string tagName, TVal value)
            where TVal : notnull
            => _contents.Add(tagName, new Content<TVal>(tagName, value));
    }
}