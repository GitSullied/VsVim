using System;
using Vim;
using System.ComponentModel.Composition;

namespace VsVim.Implementation.Ole
{
    [Export(typeof(IOleUtil))]
    internal sealed class OleUtil : IOleUtil
    {
        private bool TryConvert(CommandId commandId, IntPtr variantIn, KeyModifiers keyModifiers, out EditCommand editCommand)
        {
            return OleCommandUtil.TryConvert(commandId.Group, commandId.Id, variantIn, keyModifiers, out editCommand);
        }

        #region IOleUtil

        bool IOleUtil.TryConvert(CommandId commandId, out EditCommand editCommand)
        {
            return TryConvert(commandId, IntPtr.Zero, KeyModifiers.None, out editCommand);
        }

        bool IOleUtil.TryConvert(CommandId commandId, IntPtr variantIn, out EditCommand editCommand)
        {
            return TryConvert(commandId, variantIn, KeyModifiers.None, out editCommand);
        }

        bool IOleUtil.TryConvert(CommandId commandId, IntPtr variantIn, KeyModifiers keyModifiers, out EditCommand editCommand)
        {
            return TryConvert(commandId, variantIn, keyModifiers, out editCommand);
        }

        #endregion
    }
}
