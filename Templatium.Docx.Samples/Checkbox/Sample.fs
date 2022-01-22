module internal Templatium.Docx.Samples.Checkbox

open System.IO
open DocumentFormat.OpenXml.Packaging
open Templatium.Docx
open Templatium.Docx.Processors

[<Literal>]
let inputPath = __SOURCE_DIRECTORY__ + "/input.docx"

[<Literal>]
let outputPath = __SOURCE_DIRECTORY__ + "/output.docx"

let run () =
    let contents: IContent seq =
        [ { Title = "TurnOn"
            Value = true }
          { Title = "TurnOff"
            Value = false } ]

    use doc =
        WordprocessingDocument.Open(new MemoryStream(File.ReadAllBytes(inputPath)), true)

    doc
    |> DocxTemplater.fillDocument [ CheckBoxProcessor() ] contents
    |> DocxTemplater.deleteContentControls

    doc.SaveAs outputPath |> ignore
    ()