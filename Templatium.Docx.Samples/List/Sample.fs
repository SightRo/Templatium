module internal Templatium.Docx.Samples.List

open System.IO
open DocumentFormat.OpenXml.Packaging
open Templatium.Docx
open Templatium.Docx.Processors

[<Literal>]
let inputPath = __SOURCE_DIRECTORY__ + "/input.docx"

[<Literal>]
let outputPath = __SOURCE_DIRECTORY__ + "/output.docx"

let run () =
   
    let list: List<IContent> = [
        { Title = "Name"; Value = "Bread" }
        { Title = "Name"; Value = "Milk" }
        { Title = "Name"; Value = "Spaghetti" }
    ]

    let listContent =
        { Title = "List"
          Items = list }

    use ms = new MemoryStream()
    ms.Write(File.ReadAllBytes(inputPath))
    use doc = WordprocessingDocument.Open(ms, true)

    doc
    |> DocxTemplater.fillDocument Processors.defaults [ listContent ]
    |> DocxTemplater.deleteContentControls

    doc.SaveAs outputPath |> ignore
    ()