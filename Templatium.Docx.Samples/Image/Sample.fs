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
let image1Path = __SOURCE_DIRECTORY__ + "/image1.png"

let run () =
    use image = new FileStream(imagePath, FileMode.Open)
    use image1 = new FileStream(image1Path, FileMode.Open)

    let contents: IContent seq =
        [ { Title = "ReplaceImageWithOriginalSize"
            Image = image
            Type = Jpeg
            Format = Original
            ImagePartBehavior = Replace }
          { Title = "ReplaceImageWithExplicitSize"
            Image = image1
            Type = Png
            Format = Size { Width = 14000000; Height = 3250000 }
            ImagePartBehavior = Replace }
          // Currently doesn't work. Need really good debugging skills to fix this.
//          { Title = "AddImage"
//            Image = image1
//            Type = Png
//            Size = { Width = 12500000; Height = 5250000 } }
          ]

    use doc =
        WordprocessingDocument.Open(new MemoryStream(File.ReadAllBytes(inputPath)), true)

    doc
    |> DocxTemplater.fillDocument [ ImageProcessor() ] contents
    |> DocxTemplater.deleteContentControls

    doc.SaveAs outputPath |> ignore
    ()