using System;
using System.Collections.ObjectModel;
using EditorUtils;
using Microsoft.FSharp.Core;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.TextManager.Interop;
using Vim;
using Vim.Extensions;
using Vim.UI.Wpf;

namespace VsVim
{
    internal sealed class TransitionCommandNode : IEditCommandNode
    {
        private readonly IVimBuffer _vimBuffer;
        private readonly IVim _vim;
        private readonly IVimBufferCoordinator _bufferCoordinator;
        private readonly ITextBuffer _textBuffer;
        private readonly IVsAdapter _vsAdapter;
        private readonly IDisplayWindowBroker _broker;
        private readonly IResharperUtil _resharperUtil;
        private readonly IKeyUtil _keyUtil;

        internal TransitionCommandNode(
            IVimBufferCoordinator bufferCoordinator,
            IVsAdapter vsAdapter,
            IDisplayWindowBroker broker,
            IResharperUtil resharperUtil,
            IKeyUtil keyUtil)
        {
            _vimBuffer = bufferCoordinator.VimBuffer;
            _vim = _vimBuffer.Vim;
            _bufferCoordinator = bufferCoordinator;
            _textBuffer = _vimBuffer.TextBuffer;
            _vsAdapter = vsAdapter;
            _broker = broker;
            _resharperUtil = resharperUtil;
            _keyUtil = keyUtil;
        }

        /// <summary>
        /// Try and map a KeyInput to a single KeyInput value.  This will only succeed for KeyInput 
        /// values which have no mapping or map to a single KeyInput value
        /// </summary>
        private bool TryGetSingleMapping(KeyInput original, out KeyInput mapped)
        {
            var result = _vimBuffer.GetKeyInputMapping(original);
            if (result.IsNeedsMoreInput || result.IsRecursive)
            {
                // No single mapping
                mapped = null;
                return false;
            }

            if (result.IsMapped)
            {
                var set = ((KeyMappingResult.Mapped)result).Item;
                if (!set.IsOneKeyInput)
                {
                    mapped = null;
                    return false;
                }

                mapped = set.FirstKeyInput.Value;
                return true;
            }

            // Shouldn't get here because all cases of KeyMappingResult should be
            // handled above
            Contract.Assert(false);
            mapped = null;
            return false;
        }

        /// <summary>
        /// Is this KeyInput intended to be processed by the active display window
        /// </summary>
        private bool IsDisplayWindowKey(KeyInput keyInput)
        {
            // Consider normal completion
            if (_broker.IsCompletionActive || _resharperUtil.IsInstalled)
            {
                return
                    keyInput.IsArrowKey ||
                    keyInput == KeyInputUtil.EnterKey ||
                    keyInput == KeyInputUtil.TabKey ||
                    keyInput.Key == VimKey.Back;
            }

            if (_broker.IsSmartTagSessionActive)
            {
                return
                    keyInput.IsArrowKey ||
                    keyInput == KeyInputUtil.EnterKey;
            }

            if (_broker.IsSignatureHelpActive)
            {
                return keyInput.IsArrowKey;
            }

            return false;
        }

        /// <summary>
        /// Try and process the KeyInput from the Exec method.  This method decides whether or not
        /// a key should be processed directly by IVimBuffer or if should be going through 
        /// IOleCommandTarget.  Generally the key is processed by IVimBuffer but for many intellisense
        /// scenarios we want the key to be routed to Visual Studio directly.  Issues to consider 
        /// here are ...
        /// 
        ///  - How should the KeyInput participate in Macro playback?
        ///  - Does both VsVim and Visual Studio need to process the key (Escape mainly)
        ///  
        /// </summary>
        private bool TryProcessWithBuffer(KeyInput keyInput)
        {
            // If the IVimBuffer can't process it then it doesn't matter
            if (!_vimBuffer.CanProcess(keyInput))
            {
                return false;
            }

            // In the middle of a word completion session let insert mode handle the input.  It's 
            // displaying the intellisense itself and this method is meant to let custom intellisense
            // operate normally
            if (_vimBuffer.ModeKind == ModeKind.Insert && _vimBuffer.InsertMode.ActiveWordCompletionSession.IsSome())
            {
                return _vimBuffer.Process(keyInput).IsAnyHandled;
            }

            // The only time we actively intercept keys and route them through IOleCommandTarget
            // is when one of the IDisplayWindowBroker windows is active
            //
            // In those cases if the KeyInput is a command which should be handled by the
            // display window we route it through IOleCommandTarget to get the proper 
            // experience for those features
            if (!_broker.IsAnyDisplayActive())
            {
                // The one exception to this rule is R#.  We can't accurately determine if 
                // R# has intellisense active or not so we have to pretend like it always 
                // does.  We limit this to insert mode only though. 
                if (!_resharperUtil.IsInstalled || _vimBuffer.ModeKind != ModeKind.Insert)
                {
                    return _vimBuffer.Process(keyInput).IsAnyHandled;
                }
            }

            // Next we need to consider here are Key mappings.  The CanProcess and Process APIs 
            // will automatically map the KeyInput under the hood at the IVimBuffer level but
            // not at the individual IMode.  Have to manually map here and test against the 
            // mapped KeyInput
            KeyInput mapped;
            if (!TryGetSingleMapping(keyInput, out mapped))
            {
                return _vimBuffer.Process(keyInput).IsAnyHandled;
            }

            // If the key actually being processed is a display window key and the display window
            // is active then we allow IOleCommandTarget to control the key
            if (IsDisplayWindowKey(mapped))
            {
                return false;
            }

            var handled = _vimBuffer.Process(keyInput).IsAnyHandled;

            // The Escape key should always dismiss the active completion session.  However Vim
            // itself is mostly ignorant of display windows and typically won't dismiss them
            // as part of processing Escape (one exception is insert mode).  Dismiss it here if 
            // it's still active
            if (mapped.Key == VimKey.Escape && _broker.IsAnyDisplayActive())
            {
                _broker.DismissDisplayWindows();
            }

            return handled;
        }

        private bool Exec(EditCommand editCommand)
        {
            VimTrace.TraceInfo("VsCommandTarget::Exec {0}", editCommand);
            if (editCommand.IsUndo)
            {
                // The user hit the undo button.  Don't attempt to map anything here and instead just 
                // run a single Vim undo operation
                _vimBuffer.UndoRedoOperations.Undo(1);
                return true;
            }
            else if (editCommand.IsRedo)
            {
                // The user hit the redo button.  Don't attempt to map anything here and instead just 
                // run a single Vim redo operation
                _vimBuffer.UndoRedoOperations.Redo(1);
                return true;
            }
            else if (editCommand.HasKeyInput)
            {
                var keyInput = editCommand.KeyInput;

                // Discard the input if it's been flagged by a previous QueryStatus
                if (_bufferCoordinator.DiscardedKeyInput.IsSome(keyInput))
                {
                    return true;
                }

                // Try and process the command with the IVimBuffer
                if (TryProcessWithBuffer(keyInput))
                {
                    return true;
                }
            }

            return false;
        }

        private EditCommandStatus QueryStatus(EditCommand editCommand)
        {
            VimTrace.TraceInfo("VsCommandTarget::QueryStatus {0}", editCommand);

            _bufferCoordinator.DiscardedKeyInput = FSharpOption<KeyInput>.None;

            var action = EditCommandStatus.PassOn;
            if (editCommand.IsUndo || editCommand.IsRedo)
            {
                action = EditCommandStatus.Enable;
            }
            else if (editCommand.HasKeyInput && _vimBuffer.CanProcess(editCommand.KeyInput))
            {
                action = EditCommandStatus.Enable;
                if (_resharperUtil.IsInstalled)
                {
                    action = QueryStatusInResharper(editCommand.KeyInput) ?? EditCommandStatus.Enable;
                }
            }

            VimTrace.TraceInfo("VsCommandTarget::QueryStatus ", action);
            return action;
        }

        /// <summary>
        /// With Resharper installed we need to special certain keys like Escape.  They need to 
        /// process it in order for them to dismiss their custom intellisense but their processing 
        /// will swallow the event and not propagate it to us.  So handle, return and account 
        /// for the double stroke in exec
        /// </summary>
        private EditCommandStatus? QueryStatusInResharper(KeyInput keyInput)
        {
            EditCommandStatus? status = null;
            var passToResharper = true;
            if (_vimBuffer.ModeKind.IsAnyInsert() && keyInput == KeyInputUtil.EscapeKey)
            {
                // Have to special case Escape here for insert mode.  R# is typically ahead of us on the IOleCommandTarget
                // chain.  If a completion window is open and we wait for Exec to run R# will be ahead of us and run
                // their Exec call.  This will lead to them closing the completion window and not calling back into
                // our exec leaving us in insert mode.
                status = EditCommandStatus.Enable;
            }
            else if (_vimBuffer.ModeKind == ModeKind.ExternalEdit && keyInput == KeyInputUtil.EscapeKey)
            {
                // Have to special case Escape here for external edit mode because we want escape to get us back to 
                // normal mode.  However we do want this key to make it to R# as well since they may need to dismiss
                // intellisense
                status = EditCommandStatus.Enable;
            }
            else if ((keyInput.Key == VimKey.Back || keyInput == KeyInputUtil.EnterKey) && _vimBuffer.ModeKind != ModeKind.Insert)
            {
                // R# special cases both the Back and Enter command in various scenarios
                //
                //  - Enter is special cased in XML doc comments presumably to do custom formatting 
                //  - Enter is supressed during debugging in Exec.  Presumably this is done to avoid the annoying
                //    "Invalid ENC Edit" dialog during debugging.
                //  - Back is special cased to delete matched parens in Exec.  
                //
                // In all of these scenarios if the Enter or Back key is registered as a valid Vim
                // command we want to process it as such and prevent R# from seeing the command.  If 
                // R# is allowed to see the command they will process it often resulting in double 
                // actions
                status = EditCommandStatus.Enable;
                passToResharper = false;
            }

            // Only process the KeyInput if we are enabling the value.  When the value is Enabled
            // we return Enabled from QueryStatus and Visual Studio will push the KeyInput back
            // through the event chain where either of the following will happen 
            //
            //  1. R# will handle the KeyInput
            //  2. R# will not handle it, it will come back to use in Exec and we will ignore it
            //     because we mark it as silently handled
            if (status.HasValue && status.Value == EditCommandStatus.Enable && _vimBuffer.Process(keyInput).IsAnyHandled)
            {
                // We've broken the rules a bit by handling the command in QueryStatus and we need
                // to silently handle this command if it comes back to us again either through 
                // Exec or through the VsKeyProcessor
                _bufferCoordinator.DiscardedKeyInput = FSharpOption.Create(keyInput);

                // If we need to cooperate with R# to handle this command go ahead and pass it on 
                // to them.  Else mark it as Disabled.
                //
                // Marking it as Disabled will cause the QueryStatus call to fail.  This means the 
                // KeyInput will be routed to the KeyProcessor chain for the ITextView eventually
                // making it to our VsKeyProcessor.  That component respects the SilentlyHandled 
                // statu of KeyInput and will siently handle it
                status = passToResharper ? EditCommandStatus.Enable : EditCommandStatus.Disable;
            }

            return status;
        }

        #region IEditCommandNode

        bool IEditCommandNode.Execute(EditCommand editCommand)
        {
            return Exec(editCommand);
        }

        EditCommandStatus IEditCommandNode.QueryStatus(EditCommand editCommand)
        {
            return QueryStatus(editCommand);
        }

        #endregion
    }
}

