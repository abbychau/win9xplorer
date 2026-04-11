namespace win9xplorer
{
    internal sealed class DiagnosticsPanelForm : Form
    {
        internal sealed record DiagnosticsSnapshot(
            string ActiveHandle,
            string ForegroundHandleBeforeClick,
            int WindowCount,
            int TaskButtonCount,
            int TrayIconCount,
            int VisibleTrayIconCount,
            int ProgramCacheEntries,
            int SearchCacheEntries,
            int QuickLaunchControlCount,
            string ThemeProfile,
            string Timestamp);

        private readonly Func<DiagnosticsSnapshot> snapshotProvider;
        private readonly TextBox diagnosticsTextBox;
        private readonly System.Windows.Forms.Timer refreshTimer;

        public DiagnosticsPanelForm(Func<DiagnosticsSnapshot> snapshotProvider)
        {
            this.snapshotProvider = snapshotProvider;
            Text = "win9xplorer Diagnostics";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.SizableToolWindow;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(520, 360);

            diagnosticsTextBox = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font("Consolas", 9f, FontStyle.Regular, GraphicsUnit.Point)
            };
            Controls.Add(diagnosticsTextBox);

            refreshTimer = new System.Windows.Forms.Timer { Interval = 400 };
            refreshTimer.Tick += (_, _) => RefreshSnapshot();

            Load += (_, _) =>
            {
                RefreshSnapshot();
                refreshTimer.Start();
            };

            FormClosed += (_, _) =>
            {
                refreshTimer.Stop();
                refreshTimer.Dispose();
            };
        }

        private void RefreshSnapshot()
        {
            var snapshot = snapshotProvider();
            diagnosticsTextBox.Text =
                $"Time: {snapshot.Timestamp}{Environment.NewLine}" +
                $"Active handle: {snapshot.ActiveHandle}{Environment.NewLine}" +
                $"Foreground before click: {snapshot.ForegroundHandleBeforeClick}{Environment.NewLine}" +
                $"Windows: {snapshot.WindowCount}{Environment.NewLine}" +
                $"Task buttons: {snapshot.TaskButtonCount}{Environment.NewLine}" +
                $"Tray icons (all): {snapshot.TrayIconCount}{Environment.NewLine}" +
                $"Tray icons (visible): {snapshot.VisibleTrayIconCount}{Environment.NewLine}" +
                $"Program folder cache entries: {snapshot.ProgramCacheEntries}{Environment.NewLine}" +
                $"Program search cache entries: {snapshot.SearchCacheEntries}{Environment.NewLine}" +
                $"Quick Launch controls: {snapshot.QuickLaunchControlCount}{Environment.NewLine}" +
                $"Theme profile: {snapshot.ThemeProfile}";
        }
    }
}
