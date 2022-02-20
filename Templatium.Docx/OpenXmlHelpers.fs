namespace Templatium.Docx

open DocumentFormat.OpenXml

[<AutoOpen>]
module OpenXmlBuilder =
    type OpenXmlElement with
        member inline this.Yield(child: OpenXmlElement) =
            this.AppendChild child |> ignore
            this

        member inline this.Combine(a: OpenXmlElement, b: OpenXmlElement) : OpenXmlElement = a
        member inline this.Delay([<InlineIfLambda>] f: unit -> OpenXmlElement) = f ()

module OpenXmlHelpers =
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
