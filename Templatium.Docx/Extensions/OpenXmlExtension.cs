using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;

namespace Templatium.Docx.Extensions
{
    public static class OpenXmlExtension
    {
        public static OpenXmlElement AddAttribute(this OpenXmlElement element, OpenXmlAttribute attribute)
        {
            element.SetAttribute(attribute);
            return element;
        }

        public static IEnumerable<TElement> ChildElements<TElement>(this OpenXmlElement element)
            where TElement : OpenXmlElement
        {
            foreach (var child in element.ChildElements)
                if (child is TElement res)
                    yield return res;
        }

        public static OpenXmlElement GetFirstOrDefaultByName(this OpenXmlElement element, string name)
            => element.Descendants().Where(el => el.LocalName == name).FirstOrDefault();
        
        public static IEnumerable<OpenXmlElement> FindByName(this OpenXmlElement element, string name)
            => element.Descendants().Where(el => el.LocalName == name);
    }
}