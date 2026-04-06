namespace win9xplorer
{
    internal sealed class TaskbarOptionsForm : Form
    {
        private readonly NumericUpDown startMenuIconSizeInput;
        private readonly NumericUpDown taskIconSizeInput;
        private readonly NumericUpDown submenuOpenDelayInput;
        private readonly ComboBox fontNameCombo;
        private readonly NumericUpDown fontSizeInput;
        private readonly ComboBox buttonStyleCombo;
        private readonly NumericUpDown bevelSizeInput;
        private readonly CheckBox lazyLoadProgramsCheckBox;
        private readonly CheckBox playVolumeFeedbackSoundCheckBox;
        private readonly CheckBox autoHideTaskbarCheckBox;
        private readonly Panel taskbarBaseColorPreview;
        private readonly Panel taskbarLightColorPreview;
        private readonly Panel taskbarDarkColorPreview;
        private readonly Action openQuickLaunchFolderAction;
        private readonly Action exitAppAction;

        public event Action? ApplyRequested;

        public int StartMenuIconSize => (int)startMenuIconSizeInput.Value;

        public int TaskIconSize => (int)taskIconSizeInput.Value;

        public bool LazyLoadProgramsSubmenu => lazyLoadProgramsCheckBox.Checked;

        public bool PlayVolumeFeedbackSound => playVolumeFeedbackSoundCheckBox.Checked;

        public int StartMenuSubmenuOpenDelayMs => (int)submenuOpenDelayInput.Value;

        public bool AutoHideTaskbar => autoHideTaskbarCheckBox.Checked;

        public string TaskbarFontName => fontNameCombo.SelectedItem?.ToString() ?? "MS Sans Serif";

        public float TaskbarFontSize => (float)fontSizeInput.Value;

        public string TaskbarButtonStyle
        {
            get
            {
                var selected = buttonStyleCombo.SelectedItem?.ToString() ?? "Classic";
                return selected switch
                {
                    "Win98 Thick" => "Win98Thick",
                    _ => selected
                };
            }
        }

        public int TaskbarBevelSize => (int)bevelSizeInput.Value;

        public Color TaskbarBaseColor => taskbarBaseColorPreview.BackColor;

        public Color TaskbarLightColor => taskbarLightColorPreview.BackColor;

        public Color TaskbarDarkColor => taskbarDarkColorPreview.BackColor;

        public TaskbarOptionsForm(
            int currentStartMenuIconSize,
            int currentTaskIconSize,
            bool currentLazyLoadProgramsSubmenu,
            bool currentPlayVolumeFeedbackSound,
            int currentStartMenuSubmenuOpenDelayMs,
            bool currentAutoHideTaskbar,
            string currentTaskbarButtonStyle,
            int currentTaskbarBevelSize,
            string currentTaskbarFontName,
            float currentTaskbarFontSize,
            Color currentTaskbarBaseColor,
            Color currentTaskbarLightColor,
            Color currentTaskbarDarkColor,
            Action openQuickLaunchFolderAction,
            Action exitAppAction)
        {
            this.openQuickLaunchFolderAction = openQuickLaunchFolderAction;
            this.exitAppAction = exitAppAction;

            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            Text = "Taskbar Options";
            ClientSize = new Size(360, 474);
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

            var lblSubmenuDelay = new Label
            {
                AutoSize = true,
                Left = 12,
                Top = 72,
                Text = "Submenu open delay (ms):"
            };

            submenuOpenDelayInput = new NumericUpDown
            {
                Left = 156,
                Top = 68,
                Width = 80,
                Minimum = 0,
                Maximum = 1500,
                Increment = 25,
                Value = Math.Clamp(currentStartMenuSubmenuOpenDelayMs, 0, 1500)
            };

            var lblFontName = new Label
            {
                AutoSize = true,
                Left = 12,
                Top = 102,
                Text = "Taskbar font:"
            };

            fontNameCombo = new ComboBox
            {
                Left = 156,
                Top = 98,
                Width = 180,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            var fontNames = FontFamily.Families.Select(f => f.Name).OrderBy(n => n, StringComparer.CurrentCultureIgnoreCase).ToList();
            fontNameCombo.Items.AddRange(fontNames.Cast<object>().ToArray());
            if (fontNames.Contains(currentTaskbarFontName, StringComparer.CurrentCultureIgnoreCase))
            {
                fontNameCombo.SelectedItem = fontNames.First(n => n.Equals(currentTaskbarFontName, StringComparison.CurrentCultureIgnoreCase));
            }
            else
            {
                fontNameCombo.Items.Insert(0, currentTaskbarFontName);
                fontNameCombo.SelectedIndex = 0;
            }

            var lblFontSize = new Label
            {
                AutoSize = true,
                Left = 12,
                Top = 132,
                Text = "Taskbar font size:"
            };

            fontSizeInput = new NumericUpDown
            {
                Left = 156,
                Top = 128,
                Width = 80,
                Minimum = 7,
                Maximum = 16,
                DecimalPlaces = 2,
                Increment = 0.25m,
                Value = Math.Clamp((decimal)currentTaskbarFontSize, 7m, 16m)
            };

            var lblButtonStyle = new Label
            {
                AutoSize = true,
                Left = 12,
                Top = 162,
                Text = "Taskbar button style:"
            };

            buttonStyleCombo = new ComboBox
            {
                Left = 156,
                Top = 158,
                Width = 120,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            buttonStyleCombo.Items.AddRange(new object[] { "Classic", "Win98" });
            var styleLabel = currentTaskbarButtonStyle switch
            {
                "Win98Thick" => "Win98",
                _ => currentTaskbarButtonStyle
            };
            buttonStyleCombo.SelectedItem = buttonStyleCombo.Items.Contains(styleLabel)
                ? styleLabel
                : "Classic";

            var lblBevelSize = new Label
            {
                AutoSize = true,
                Left = 12,
                Top = 192,
                Text = "Bevel size (px):"
            };

            bevelSizeInput = new NumericUpDown
            {
                Left = 156,
                Top = 188,
                Width = 80,
                Minimum = 1,
                Maximum = 4,
                Value = Math.Clamp(currentTaskbarBevelSize, 1, 4)
            };

            var lblTaskbarBaseColor = new Label
            {
                AutoSize = true,
                Left = 12,
                Top = 222,
                Text = "Taskbar color:"
            };
            taskbarBaseColorPreview = CreateColorPreview(currentTaskbarBaseColor, left: 156, top: 218);
            var btnTaskbarBaseColor = CreateColorPickButton("...", left: 244, top: 216, onClick: () => PickColor(taskbarBaseColorPreview));
            var btnTaskbarBaseColorDefault = CreateColorPickButton("Default", left: 278, top: 216, onClick: () => taskbarBaseColorPreview.BackColor = Color.FromArgb(192, 192, 192));
            btnTaskbarBaseColorDefault.Width = 70;

            var lblTaskbarLightColor = new Label
            {
                AutoSize = true,
                Left = 12,
                Top = 252,
                Text = "Light color:"
            };
            taskbarLightColorPreview = CreateColorPreview(currentTaskbarLightColor, left: 156, top: 248);
            var btnTaskbarLightColor = CreateColorPickButton("...", left: 244, top: 246, onClick: () => PickColor(taskbarLightColorPreview));
            var btnTaskbarLightColorDefault = CreateColorPickButton("Default", left: 278, top: 246, onClick: () => taskbarLightColorPreview.BackColor = Color.FromArgb(255, 255, 255));
            btnTaskbarLightColorDefault.Width = 70;

            var lblTaskbarDarkColor = new Label
            {
                AutoSize = true,
                Left = 12,
                Top = 282,
                Text = "Dark color:"
            };
            taskbarDarkColorPreview = CreateColorPreview(currentTaskbarDarkColor, left: 156, top: 278);
            var btnTaskbarDarkColor = CreateColorPickButton("...", left: 244, top: 276, onClick: () => PickColor(taskbarDarkColorPreview));
            var btnTaskbarDarkColorDefault = CreateColorPickButton("Default", left: 278, top: 276, onClick: () => taskbarDarkColorPreview.BackColor = Color.FromArgb(128, 128, 128));
            btnTaskbarDarkColorDefault.Width = 70;

            lazyLoadProgramsCheckBox = new CheckBox
            {
                Left = 12,
                Top = 312,
                Width = 336,
                Text = "Lazy-load Programs submenu",
                Checked = currentLazyLoadProgramsSubmenu
            };

            var lazyLoadHintLabel = new Label
            {
                Left = 28,
                Top = 334,
                Width = 320,
                Height = 24,
                Text = "On: faster Start open; Off: faster Programs expand"
            };

            playVolumeFeedbackSoundCheckBox = new CheckBox
            {
                Left = 12,
                Top = 360,
                Width = 336,
                Text = "Play feedback sound on volume slider release",
                Checked = currentPlayVolumeFeedbackSound
            };

            autoHideTaskbarCheckBox = new CheckBox
            {
                Left = 12,
                Top = 388,
                Width = 336,
                Text = "Auto-hide taskbar",
                Checked = currentAutoHideTaskbar
            };

            var openQuickLaunchFolderButton = new Button
            {
                Left = 12,
                Top = 416,
                Width = 236,
                Height = 24,
                Text = "Open Quick Launch Folder..."
            };
            openQuickLaunchFolderButton.Click += (_, _) => this.openQuickLaunchFolderAction();

            var btnExitApp = new Button
            {
                Text = "Exit App",
                Left = 258,
                Top = 416,
                Width = 90,
                Height = 24
            };
            btnExitApp.Click += (_, _) => this.exitAppAction();

            var btnOk = new Button
            {
                Text = "OK",
                Left = 107,
                Top = 444,
                Width = 75,
                DialogResult = DialogResult.OK
            };

            var btnCancel = new Button
            {
                Text = "Cancel",
                Left = 188,
                Top = 444,
                Width = 75,
                DialogResult = DialogResult.Cancel
            };

            var btnApply = new Button
            {
                Text = "Apply",
                Left = 269,
                Top = 444,
                Width = 75
            };
            btnApply.Click += (_, _) => ApplyRequested?.Invoke();

            Controls.Add(lblStartMenuIcon);
            Controls.Add(startMenuIconSizeInput);
            Controls.Add(lblTaskIcon);
            Controls.Add(taskIconSizeInput);
            Controls.Add(lblSubmenuDelay);
            Controls.Add(submenuOpenDelayInput);
            Controls.Add(lblFontName);
            Controls.Add(fontNameCombo);
            Controls.Add(lblFontSize);
            Controls.Add(fontSizeInput);
            Controls.Add(lblButtonStyle);
            Controls.Add(buttonStyleCombo);
            Controls.Add(lblBevelSize);
            Controls.Add(bevelSizeInput);
            Controls.Add(lblTaskbarBaseColor);
            Controls.Add(taskbarBaseColorPreview);
            Controls.Add(btnTaskbarBaseColor);
            Controls.Add(btnTaskbarBaseColorDefault);
            Controls.Add(lblTaskbarLightColor);
            Controls.Add(taskbarLightColorPreview);
            Controls.Add(btnTaskbarLightColor);
            Controls.Add(btnTaskbarLightColorDefault);
            Controls.Add(lblTaskbarDarkColor);
            Controls.Add(taskbarDarkColorPreview);
            Controls.Add(btnTaskbarDarkColor);
            Controls.Add(btnTaskbarDarkColorDefault);
            Controls.Add(lazyLoadProgramsCheckBox);
            Controls.Add(lazyLoadHintLabel);
            Controls.Add(playVolumeFeedbackSoundCheckBox);
            Controls.Add(autoHideTaskbarCheckBox);
            Controls.Add(openQuickLaunchFolderButton);
            Controls.Add(btnExitApp);
            Controls.Add(btnOk);
            Controls.Add(btnCancel);
            Controls.Add(btnApply);

            AcceptButton = btnOk;
            CancelButton = btnCancel;

            Shown += (_, _) =>
            {
                var desiredTop = Top - 100;
                var workingArea = Screen.FromPoint(Location).WorkingArea;
                if (desiredTop < workingArea.Top)
                {
                    desiredTop = workingArea.Top;
                }

                Top = desiredTop;
            };
        }

        private static Panel CreateColorPreview(Color color, int left, int top)
        {
            return new Panel
            {
                Left = left,
                Top = top,
                Width = 80,
                Height = 22,
                BackColor = color,
                BorderStyle = BorderStyle.FixedSingle
            };
        }

        private static Button CreateColorPickButton(string text, int left, int top, Action onClick)
        {
            var button = new Button
            {
                Text = text,
                Left = left,
                Top = top,
                Width = 30,
                Height = 24
            };
            button.Click += (_, _) => onClick();
            return button;
        }

        private void PickColor(Panel preview)
        {
            using var dialog = new ColorDialog
            {
                FullOpen = true,
                Color = preview.BackColor
            };

            if (dialog.ShowDialog(this) == DialogResult.OK)
            {
                preview.BackColor = dialog.Color;
            }
        }
    }
}
