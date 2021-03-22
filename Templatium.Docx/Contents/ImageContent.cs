namespace Templatium.Docx.Contents
{
    public class ImageContent : IContent
    {
        public ImageContent(string tagName, byte[] value, string extension)
        {
            TagName = tagName;
            Value = value;
            Extension = extension;
        }

        public string TagName { get; init; }
        public int? Width { get; init; }
        public int? Height { get; init; }
        public string Extension { get; init; }
        public byte[] Value { get; init; }
        object IContent.Value => Value;
    }
}