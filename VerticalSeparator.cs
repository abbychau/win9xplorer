using System.Drawing;
using System.Windows.Forms;

namespace win9xplorer
{
    internal sealed class VerticalSeparator : Panel
    {
        private Color separatorDarkColor;
        private Color separatorLightColor;

        public VerticalSeparator(Color backgroundColor, Color darkColor, Color lightColor)
        {
            separatorDarkColor = darkColor;
            separatorLightColor = lightColor;
            Dock = DockStyle.Left;
            Width = 4;
            Margin = Padding.Empty;
            Padding = Padding.Empty;
            BackColor = backgroundColor;
            Paint += VerticalSeparator_Paint;
        }

        public void ApplyTheme(Color backgroundColor, Color darkColor, Color lightColor)
        {
            BackColor = backgroundColor;
            separatorDarkColor = darkColor;
            separatorLightColor = lightColor;
            Invalidate();
        }

        private void VerticalSeparator_Paint(object? sender, PaintEventArgs e)
        {
            if (sender is not Panel panel)
            {
                return;
            }

            e.Graphics.Clear(panel.BackColor);
            var center = panel.Width / 2;
            var top = 0;
            var bottom = Math.Max(top, panel.Height - 3);
            using var darkPen = new Pen(separatorDarkColor);
            using var lightPen = new Pen(separatorLightColor);
            e.Graphics.DrawLine(darkPen, center - 1, top, center - 1, bottom);
            e.Graphics.DrawLine(lightPen, center, top, center, bottom);
        }
    }
}
