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
[<Literal>]
let image2Path = __SOURCE_DIRECTORY__ + "/image2.jpg"

let run () =
    use image = new FileStream(imagePath, FileMode.Open)
    use image1 = new FileStream(image1Path, FileMode.Open)
    use image2 = new FileStream(image2Path, FileMode.Open)

    let contents: IContent seq =
        [ { Title = "ReplaceImageWithOriginalSize"
            Image = image
            Type = Jpeg
            Format = Original
            ImagePartBehavior = Replace }
          { Title = "ReplaceImageWithExplicitSize"
            Image = image1
            Type = Png
            Format = Size { Width = 5200000; Height = 2500000 }
            ImagePartBehavior = Replace }
          { Title = "AddImage"
            Image = image2
            Type = Jpeg
            Size = { Width = 3000000; Height = 2000000 } }
          ]

    use doc =
        WordprocessingDocument.Open(new MemoryStream(File.ReadAllBytes(inputPath)), true)

    doc
    |> DocxTemplater.fillDocument [ ImageProcessor() ] contents
    |> DocxTemplater.deleteContentControls

    doc.SaveAs outputPath |> ignore
    ()