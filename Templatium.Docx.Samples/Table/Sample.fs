module internal Templatium.Docx.Samples.Table

open System.IO
open DocumentFormat.OpenXml.Packaging
open Templatium.Docx
open Templatium.Docx.Processors

[<Literal>]
let inputPath = __SOURCE_DIRECTORY__ + "/input.docx"

[<Literal>]
let outputPath = __SOURCE_DIRECTORY__ + "/output.docx"

[<Literal>]
let facebook = __SOURCE_DIRECTORY__ + "/facebook.png"

[<Literal>]
let netscape = __SOURCE_DIRECTORY__ + "/netscape.png"

let run () =
    let rows: List<List<IContent>> =
        [ [ { Title = "Name"; Value = "Facebook" }
            { Title = "Logo"
              Image = File.ReadAllBytes(facebook)
              Type = Png
              Size = Original }
            { Title = "IsActive"; Value = true } ]
          [ { Title = "Name"; Value = "Netscape" }
            { Title = "Logo"
              Image = File.ReadAllBytes(netscape)
              Type = Png
              Size = Original }
            { Title = "IsActive"; Value = false } ] ]

    let tableContent =
        { Title = "CompaniesTable"
          Rows = rows }

    use ms = new MemoryStream()
    ms.Write(File.ReadAllBytes(inputPath))
    use doc = WordprocessingDocument.Open(ms, true)

    doc
    |> DocxTemplater.fillDocument Processors.defaults [ tableContent ]
    |> DocxTemplater.deleteContentControls

    doc.SaveAs outputPath |> ignore
    ()