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
let imagePath = __SOURCE_DIRECTORY__ + "/image.jpg"
[<Literal>]
let image1Path = __SOURCE_DIRECTORY__ + "/image.jpg"

let run () =
    let contents: IContent seq =
        [ { Title = "ReplaceImageWithOriginalSize"
            Image = File.ReadAllBytes imagePath
            Type = Jpeg
            Size = Original }
          { Title = "ReplaceImageWithExplicitSize"
            Image = File.ReadAllBytes image1Path
            Type = Jpeg
            Size = Size(width = 14000000, height = 3250000) }
          // Currently doesn't work. Need really good debugging skills to fix this. 
//          { Title = "AddImage"
//            Image = imageBytes
//            Type = Jpeg
//            Size = Size(width = 12500000, height = 5250000) }
          ]

    use doc =
        WordprocessingDocument.Open(new MemoryStream(File.ReadAllBytes(inputPath)), true)

    doc
    |> DocxTemplater.fillDocument [ ImageProcessor() ] contents
    |> DocxTemplater.deleteContentControls

    doc.SaveAs outputPath |> ignore
    ()