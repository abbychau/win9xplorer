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
        public ToolbarComponents(Panel panel, Func<Control> gripFactory)
        {
            Panel = panel;
            Grip = gripFactory();
        }

        public Panel Panel { get; }

        public Control Grip { get; }
    }
}
