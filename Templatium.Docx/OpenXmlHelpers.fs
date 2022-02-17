namespace Templatium.Docx

open DocumentFormat.OpenXml

module OpenXmlHelpers =

    type OpenXmlElement with
        member this.With children =
            this.Append(children)
            this

    let findFirstNodeByName<'t when 't :> OpenXmlElement> (element: OpenXmlElement) name =
        element.Descendants()
        |> Seq.filter (fun el -> el.LocalName = name)
        |> Seq.tryHead
        |> Option.bind
            (fun el ->
                match el with
                | :? 't as res -> Some res
                | _ -> None)

    let findDescendantsByName<'t when 't :> OpenXmlElement> (element: OpenXmlElement) name =
        element.Descendants()
        |> Seq.filter (fun el -> el.LocalName = name)
        |> Seq.choose
            (fun el ->
                match el with
                | :? 't as res -> Some res
                | _ -> None)
