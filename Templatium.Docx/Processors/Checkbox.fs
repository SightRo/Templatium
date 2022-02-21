namespace Templatium.Docx.Processors

open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Office2010.Word
open Templatium.Docx
open DocumentFormat.OpenXml.Wordprocessing

module private ProcessorImpl =
    [<Literal>]
    let checkedUnicode = 0x2612

    [<Literal>]
    let checkedSymbol = "☒"

    [<Literal>]
    let uncheckedUnicode = 0x2610

    [<Literal>]
    let uncheckedSymbol = "☐"

    let addCheckbox (sdt: SdtElement) content =
        let symbol =
            if content.Value then
                checkedSymbol
            else
                uncheckedSymbol

        let checkedNode = Checked()
        checkedNode.Val <- OnOffValue(content.Value)
        let checkedStateNode = CheckedState()
        checkedStateNode.Val <- HexBinaryValue(string checkedUnicode)
        let uncheckedStateNode = UncheckedState()
        uncheckedStateNode.Val <- HexBinaryValue(string uncheckedUnicode)

        sdt.RemoveAllChildren()

        sdt.SdtProperties.AppendChild(
            SdtContentCheckBox() {
                checkedNode
                checkedStateNode
                uncheckedStateNode
            }
        )
        |> ignore

        sdt
            .GetFirstChild<SdtContentRun>()
            .AppendChild(Run() { Text(symbol) })
        |> ignore

        ()

    let updateCheckbox (sdt: SdtElement) (checkboxNode: SdtContentCheckBox) content =
        let symbol =
            if content.Value then
                checkedSymbol
            else
                uncheckedSymbol

        checkboxNode.Checked.Val.Value <-
            (if content.Value then
                 OnOffValues.One
             else
                 OnOffValues.Zero)

        OpenXmlHelpers.findFirstNodeByName<Text> sdt Constants.text
        |> Option.iter (fun textNode -> textNode.Text <- symbol)

type CheckBoxProcessor() =
    interface IProcessor with
        member _.CanFill content _ _ = content :? Content<bool>

        member _.Fill content sdt _ =
            let boolContent = content :?> Content<bool>

            let checkboxNodeOpt =
                OpenXmlHelpers.findFirstNodeByName<SdtContentCheckBox> sdt Constants.checkbox

            match checkboxNodeOpt with
            | Some checkboxNode -> ProcessorImpl.updateCheckbox sdt checkboxNode boolContent
            | _ -> ProcessorImpl.addCheckbox sdt boolContent

            ()
