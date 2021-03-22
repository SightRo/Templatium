namespace Templatium.Docx.Contents
{
    public interface IContent
    {
        string TagName { get; }
        object Value { get; }
    }
}