
using System;
using Vim;
namespace VsVim
{
    public interface IOleUtil
    {
        /// <summary>
        /// Try and convert the given OleCommandData into an actionable command
        /// </summary>
        bool TryConvert(CommandId commandId, out EditCommand editCommand);

        /// <summary>
        /// Try and convert the given OleCommandData into an actionable command
        /// </summary>
        bool TryConvert(CommandId commandId, IntPtr variantIn, out EditCommand editCommand);

        /// <summary>
        /// Try and convert the given OleCommandData into an actionable command pretending
        /// the given KeyModifiers were pressed
        /// </summary>
        bool TryConvert(CommandId commandId, IntPtr variantIn, KeyModifiers keyModifiers, out EditCommand editCommand);
    }
}
