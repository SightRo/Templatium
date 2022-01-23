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
let facebookPath = __SOURCE_DIRECTORY__ + "/facebook.png"

[<Literal>]
let netscapePath = __SOURCE_DIRECTORY__ + "/netscape.png"

let run () =
    use facebook =
        new FileStream(facebookPath, FileMode.Open)

    use netscape =
        new FileStream(netscapePath, FileMode.Open)

    let rows: List<List<IContent>> =
        [ [ { Title = "Name"; Value = "Facebook" }
            { Title = "Logo"
              Image = facebook
              Type = Png
              Format = Original
              ImagePartBehavior = Add }
            { Title = "IsActive"; Value = true } ]
          [ { Title = "Name"; Value = "Netscape" }
            { Title = "Logo"
              Image = netscape
              Type = Png
              Format = Original
              ImagePartBehavior = Add }
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