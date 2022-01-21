namespace Templatium.Docx.Processors

open Templatium.Docx
open Templatium.Docx.Processors

[<RequireQualifiedAccess>]
module Processors =
    let defaults: IProcessor seq =
        [ StringProcessor()
          ImageProcessor()
          TableProcessor()
          CheckBoxProcessor() ]