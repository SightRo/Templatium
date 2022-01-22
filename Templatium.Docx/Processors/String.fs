namespace Templatium.Docx.Processors

open Templatium.Docx
open DocumentFormat.OpenXml.Wordprocessing
open System.Linq

type StringProcessor() =
    interface IProcessor with
        member _.CanFill content _ _ = content :? Content<string>

        member _.Fill content sdt _ =
            let stringContent = content :?> Content<string>

            let contentNode =
                OpenXmlHelpers.findFirstNodeByName sdt Constants.sdtContent

            match contentNode with
            | Some block ->
                let textNodes = block.Descendants<Text>().ToList()

                match textNodes.Count with
                | 0 ->
                    // This block will never executes if document created using Word
                    // At least the latest Word insert placeholder text in any case
                    // So it not tested in any sense
                    block.RemoveAllChildren()

                    block.AppendChild(Paragraph(Run(Text(stringContent.Value))))
                    |> ignore
                | _ ->
                    // If text is multiline, Word insert multiple text nodes (one per line)
                    // So we modify the first one to preserve formatting and delete the rest
                    textNodes[0].Text <- stringContent.Value

                    textNodes
                    |> Seq.skip 1
                    |> Seq.iter (fun n -> n.Remove())
            | _ -> ()

            ()
