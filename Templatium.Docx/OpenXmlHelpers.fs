namespace Templatium.Docx

open DocumentFormat.OpenXml
open System.Linq

module OpenXmlHelpers =
    let findFirstNodeByName<'t when 't :> OpenXmlElement> (element: OpenXmlElement) name =
        let res =
            element
                .Descendants()
                .Where(fun d -> d.LocalName = name)
                .FirstOrDefault()

        match res with
        | null -> None
        | node ->
            match node with
                | :? 't as res -> Some res
                | _ -> None
                
    let findAllNodeByName<'t when 't :> OpenXmlElement> (element: OpenXmlElement) name =
        element.Descendants().Where(fun e -> e.LocalName = name).OfType<'t>()