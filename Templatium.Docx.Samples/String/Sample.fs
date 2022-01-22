module internal Templatium.Docx.Samples.String

open System.IO
open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open Templatium.Docx
open Templatium.Docx.Processors

[<Literal>]
let inputPath = __SOURCE_DIRECTORY__ + "/input.docx"

[<Literal>]
let outputPath = __SOURCE_DIRECTORY__ + "/output.docx"

let run () =
    let contents: IContent seq =
        [ { Title = "ReplaceText"
            Value = "This text has been replaced and original formatting was kept" }
          { Title = "AddText"
            Value = "This text was added automatically without any formatting" } ]

    use doc =
        WordprocessingDocument.Open(new MemoryStream(File.ReadAllBytes(inputPath)), true)

    doc
    |> DocxTemplater.fillDocument [ StringProcessor() ] contents
    |> DocxTemplater.deleteContentControls

    doc.SaveAs outputPath |> ignore
    ()