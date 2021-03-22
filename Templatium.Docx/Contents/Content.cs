namespace Templatium.Docx.Contents
{
    public class Content<TVal> : IContent
        where TVal : notnull
    {
        public Content(string tagName, TVal value)
        {
            TagName = tagName;
            Value = value;
        }

        public string TagName { get; init; }
        public TVal Value { get; init; }
        object IContent.Value => Value;
    }
}