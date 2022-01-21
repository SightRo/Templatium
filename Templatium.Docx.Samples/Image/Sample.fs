module internal Templatium.Docx.Samples.Image

open System.IO
open DocumentFormat.OpenXml.Packaging
open Templatium.Docx
open Templatium.Docx.Processors

[<Literal>]
let inputPath = __SOURCE_DIRECTORY__ + "/input.docx"

[<Literal>]
let outputPath = __SOURCE_DIRECTORY__ + "/output.docx"

[<Literal>]
let imagePath = __SOURCE_DIRECTORY__ + "/image.png"

let run () =
    let imageBytes = File.ReadAllBytes imagePath

    let contents: IContent seq =
        [ { Title = "ReplaceImage"
            Image = imageBytes
            Type = Png
            Size = Original }
          { Title = "AddImage"
            Image = imageBytes
            Type = Png
            Size = Size(width = 1250, height = 525) } ]

    use doc =
        WordprocessingDocument.Open(inputPath, true)

    doc
    |> DocxTemplater.fillDocument [ImageProcessor()] contents
    |> DocxTemplater.deleteContentControls

    doc.SaveAs outputPath |> ignore
    ()