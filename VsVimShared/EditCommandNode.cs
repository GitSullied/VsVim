
namespace VsVim
{
    internal enum EditCommandStatus
    {
        /// <summary>
        /// Command is enabled
        /// </summary>
        Enable,

        /// <summary>
        /// Command is disabled
        /// </summary>
        Disable,

        /// <summary>
        /// VsVim isn't concerned about the command and it's left to the next IOleCommandTarget
        /// to determine if it's enabled or not
        /// </summary>
        PassOn,
    }

    internal interface IEditCommandNode
    {
        bool Execute(EditCommand editCommand);

        EditCommandStatus QueryStatus(EditCommand editCommand);
    }
}
