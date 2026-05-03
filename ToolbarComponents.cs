using System.Windows.Forms;

namespace win9xplorer
{
    internal enum ToolbarId
    {
        QuickLaunch,
        Address,
        Volume,
        Spotify
    }

    internal sealed class ToolbarComponents
    {
        public ToolbarComponents(Panel panel, Control grip)
        {
            Panel = panel;
            Grip = grip;
        }

        public Panel Panel { get; }

        public Control Grip { get; }
    }
}
