namespace Templatium.Docx.Processors

open Templatium.Docx
open DocumentFormat.OpenXml.Wordprocessing
open System.Linq

type StringProcessor =
    interface IProcessor with
        member _.CanFill _ _ content = content :? Content<string>

        member _.Fill _ sdt content =
            let stringContent = content :?> Content<string>

            let contentNode =
                OpenXmlHelpers.findFirstNodeByName sdt Constants.sdtContent

            match contentNode with
            | Some block ->
                let textNode =
                    block.Descendants<Text>().FirstOrDefault()

                match textNode with
                | null ->
                    block.RemoveAllChildren()

                    block.AppendChild(Paragraph(Run(Text(stringContent.Value))))
                    |> ignore
                | _ -> textNode.Text <- stringContent.Value
            | _ -> ()

            ()
