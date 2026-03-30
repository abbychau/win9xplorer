namespace win9xplorer
{
    internal sealed class TimeDetailsForm : Form
    {
        private readonly Label timeLabel;
        private readonly Label dateLabel;
        private readonly System.Windows.Forms.Timer tickTimer;

        public TimeDetailsForm()
        {
            FormBorderStyle = FormBorderStyle.FixedToolWindow;
            StartPosition = FormStartPosition.Manual;
            ShowInTaskbar = false;
            TopMost = true;
            MaximizeBox = false;
            MinimizeBox = false;
            Text = "Date and Time";
            ClientSize = new Size(250, 116);
            Font = new Font("MS Sans Serif", 8.25f, FontStyle.Regular, GraphicsUnit.Point);

            timeLabel = new Label
            {
                Left = 16,
                Top = 16,
                Width = 218,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("MS Sans Serif", 12f, FontStyle.Bold, GraphicsUnit.Point)
            };

            dateLabel = new Label
            {
                Left = 16,
                Top = 52,
                Width = 218,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = Font
            };

            var closeButton = new Button
            {
                Left = 160,
                Top = 82,
                Width = 74,
                Height = 24,
                Text = "Close",
                DialogResult = DialogResult.OK
            };

            Controls.Add(timeLabel);
            Controls.Add(dateLabel);
            Controls.Add(closeButton);

            AcceptButton = closeButton;
            CancelButton = closeButton;

            tickTimer = new System.Windows.Forms.Timer { Interval = 1000 };
            tickTimer.Tick += (_, _) => UpdateDateTime();

            VisibleChanged += (_, _) =>
            {
                if (Visible)
                {
                    UpdateDateTime();
                    tickTimer.Start();
                }
                else
                {
                    tickTimer.Stop();
                }
            };

            Deactivate += (_, _) => Hide();

            FormClosed += (_, _) =>
            {
                tickTimer.Stop();
                tickTimer.Dispose();
                timeLabel.Font.Dispose();
            };
        }

        public void ShowAt(Point location)
        {
            Location = location;
            UpdateDateTime();
            Show();
            Activate();
        }

        private void UpdateDateTime()
        {
            var now = DateTime.Now;
            timeLabel.Text = now.ToString("hh:mm:ss tt");
            dateLabel.Text = now.ToLongDateString();
        }
    }
}
