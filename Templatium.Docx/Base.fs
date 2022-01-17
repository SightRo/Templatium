namespace Templatium.Docx

open DocumentFormat.OpenXml.Wordprocessing
open DocumentFormat.OpenXml.Packaging

type IContent =
    abstract Title : string
    abstract Value : obj

type Content<'a> =
    { Title: string
      Value: 'a }
    interface IContent with
        member this.Title = this.Title
        member this.Value = this.Value :> obj

type IProcessor =
    abstract CanFill: WordprocessingDocument -> SdtElement -> IContent -> bool
    abstract Fill: WordprocessingDocument -> SdtElement -> IContent -> unit


