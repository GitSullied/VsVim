﻿#light

namespace Vim.Modes.Insert
open Vim
open Microsoft.VisualStudio.Text
open Microsoft.VisualStudio.Text.Operations
open Microsoft.VisualStudio.Text.Editor
open System.Windows.Input
open System.Windows.Media

type internal DefaultOperations(_textView : ITextView, _editorOpts : IEditorOperations ) =
    inherit Modes.CommonOperations(_textView, _editorOpts)