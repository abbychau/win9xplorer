namespace win9xplorer
{
    internal sealed class TaskbarOptionsForm : Form
    {
        private readonly NumericUpDown startMenuIconSizeInput;
        private readonly NumericUpDown taskIconSizeInput;
        private readonly CheckBox lazyLoadProgramsCheckBox;
        private readonly Action openQuickLaunchFolderAction;

        public int StartMenuIconSize => (int)startMenuIconSizeInput.Value;

        public int TaskIconSize => (int)taskIconSizeInput.Value;

        public bool LazyLoadProgramsSubmenu => lazyLoadProgramsCheckBox.Checked;

        public TaskbarOptionsForm(
            int currentStartMenuIconSize,
            int currentTaskIconSize,
            bool currentLazyLoadProgramsSubmenu,
            Action openQuickLaunchFolderAction)
        {
            this.openQuickLaunchFolderAction = openQuickLaunchFolderAction;

            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            Text = "Taskbar Options";
            ClientSize = new Size(300, 202);
            Font = new Font("MS Sans Serif", 8.25f, FontStyle.Regular, GraphicsUnit.Point);

            var lblStartMenuIcon = new Label
            {
                AutoSize = true,
                Left = 12,
                Top = 14,
                Text = "Start/menu icon size:"
            };

            startMenuIconSizeInput = new NumericUpDown
            {
                Left = 156,
                Top = 10,
                Width = 80,
                Minimum = 16,
                Maximum = 32,
                Value = Math.Clamp(currentStartMenuIconSize, 16, 32)
            };

            var lblTaskIcon = new Label
            {
                AutoSize = true,
                Left = 12,
                Top = 44,
                Text = "Task icon size:"
            };

            taskIconSizeInput = new NumericUpDown
            {
                Left = 156,
                Top = 40,
                Width = 80,
                Minimum = 16,
                Maximum = 32,
                Value = Math.Clamp(currentTaskIconSize, 16, 32)
            };

            lazyLoadProgramsCheckBox = new CheckBox
            {
                Left = 12,
                Top = 72,
                Width = 272,
                Text = "Lazy-load Programs submenu",
                Checked = currentLazyLoadProgramsSubmenu
            };

            var openQuickLaunchFolderButton = new Button
            {
                Left = 12,
                Top = 100,
                Width = 272,
                Height = 24,
                Text = "Open Quick Launch Folder..."
            };
            openQuickLaunchFolderButton.Click += (_, _) => this.openQuickLaunchFolderAction();

            var btnOk = new Button
            {
                Text = "OK",
                Left = 116,
                Top = 156,
                Width = 75,
                DialogResult = DialogResult.OK
            };

            var btnCancel = new Button
            {
                Text = "Cancel",
                Left = 196,
                Top = 156,
                Width = 75,
                DialogResult = DialogResult.Cancel
            };

            Controls.Add(lblStartMenuIcon);
            Controls.Add(startMenuIconSizeInput);
            Controls.Add(lblTaskIcon);
            Controls.Add(taskIconSizeInput);
            Controls.Add(lazyLoadProgramsCheckBox);
            Controls.Add(openQuickLaunchFolderButton);
            Controls.Add(btnOk);
            Controls.Add(btnCancel);

            AcceptButton = btnOk;
            CancelButton = btnCancel;
        }
    }
}
