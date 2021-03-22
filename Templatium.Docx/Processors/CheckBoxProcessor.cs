using System;
using System.Linq;
using DocumentFormat.OpenXml.Office2010.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Templatium.Docx.Contents;
using Templatium.Docx.Extensions;
using Checked = DocumentFormat.OpenXml.Office2010.Word.Checked;

namespace Templatium.Docx.Processors
{
    public class CheckBoxProcessor : IProcessor
    {
        private const string DefaultCheckedStateUnicodeId = "2612";
        private const string DefaultUncheckedStateUnicodeId = "2610";
        private const string DefaultCheckedStateSymbol = "☒";
        private const string DefaultUncheckedStateSymbol = "☐";

        public Type TargetType => typeof(Content<bool>);

        public bool CanFill(IDocxTemplater templater, WordprocessingDocument document, SdtElement sdtElement,
            IContent content)
            => content is Content<bool>;

        public void FillContent(IDocxTemplater templater, WordprocessingDocument document, SdtElement sdtElement,
            IContent content)
        {
            var boolContent = (Content<bool>) content;

            var checkBox = sdtElement.SdtProperties.GetFirstChild<SdtContentCheckBox>();
            if (checkBox != null)
                UpdateCheckBox(sdtElement, boolContent, checkBox);
            else
                AddCheckBox(sdtElement, boolContent);
        }

        private void AddCheckBox(SdtElement sdtElement, Content<bool> content)
        {
            var symbol = content.Value ? DefaultCheckedStateSymbol : DefaultUncheckedStateSymbol;
            var value = content.Value ? "1" : "0";

            var checkedElement = new Checked()
                .AddAttribute(new("w14", "val", "http://schemas.microsoft.com/office/word/2010/wordml", value));
            var checkedState = new CheckedState()
                .AddAttribute(new("w14", "val", "http://schemas.microsoft.com/office/word/2010/wordml",
                    DefaultCheckedStateUnicodeId));
            var uncheckedState = new UncheckedState()
                .AddAttribute(new("w14", "val", "http://schemas.microsoft.com/office/word/2010/wordml",
                    DefaultUncheckedStateUnicodeId));

            sdtElement.SdtProperties.AppendChild(new SdtContentCheckBox(
                checkedElement,
                checkedState,
                uncheckedState
            ));

            sdtElement.GetFirstChild<SdtContentRun>().AppendChild(
                new Run(
                    new Text(symbol)
                )
            );
        }

        private void UpdateCheckBox(SdtElement sdtElement, Content<bool> boolContent, SdtContentCheckBox checkBox)
        {
            var selectedSymbol = boolContent.Value
                ? GetUnicodeSymbol(checkBox.CheckedState.Val.Value)
                : GetUnicodeSymbol(checkBox.UncheckedState.Val.Value);

            checkBox.Checked.Val.Value = boolContent.Value ? OnOffValues.One : OnOffValues.Zero;
            var textElement = sdtElement.GetFirstChild<SdtContentRun>().Descendants<Text>().FirstOrDefault();
            textElement.Text = selectedSymbol;
        }

        private string GetUnicodeSymbol(string code)
            => char.ConvertFromUtf32(Convert.ToInt32(code, 16));
    }
}