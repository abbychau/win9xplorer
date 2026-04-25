using System.Diagnostics;
using System.Text.Json;

namespace win9xplorer
{
    internal sealed class TaskbarOptionsForm : Form
    {
        private sealed record ColorRowControls(Panel Preview, TextBox HexBox);
        private sealed record ThemeClipboardPayload(
            string ThemeProfileName,
            string TaskbarButtonStyle,
            int TaskbarBevelSize,
            string TaskbarFontName,
            float TaskbarFontSize,
            string TaskbarFontColor,
            string TaskbarBaseColor,
            string TaskbarLightColor,
            string TaskbarDarkColor);

        private readonly NumericUpDown startMenuIconSizeInput;
        private readonly NumericUpDown taskIconSizeInput;
        private readonly NumericUpDown submenuOpenDelayInput;
        private readonly ComboBox fontNameCombo;
        private readonly NumericUpDown fontSizeInput;
        private readonly ComboBox buttonStyleCombo;
        private readonly NumericUpDown bevelSizeInput;
        private readonly CheckBox lazyLoadProgramsCheckBox;
        private readonly CheckBox playVolumeFeedbackSoundCheckBox;
        private readonly CheckBox useClassicVolumePopupCheckBox;
        private readonly CheckBox autoHideTaskbarCheckBox;
        private readonly CheckBox startOnWindowsStartupCheckBox;
        private readonly ComboBox themeProfileCombo;
        private readonly Action openQuickLaunchFolderAction;

        private readonly ColorRowControls taskbarBaseColorRow;
        private readonly ColorRowControls taskbarLightColorRow;
        private readonly ColorRowControls taskbarDarkColorRow;
        private readonly ColorRowControls taskbarFontColorRow;

        public event Action? ApplyRequested;

        public int StartMenuIconSize => (int)startMenuIconSizeInput.Value;
        public int TaskIconSize => (int)taskIconSizeInput.Value;
        public bool LazyLoadProgramsSubmenu => lazyLoadProgramsCheckBox.Checked;
        public bool PlayVolumeFeedbackSound => playVolumeFeedbackSoundCheckBox.Checked;
        public bool UseClassicVolumePopup => useClassicVolumePopupCheckBox.Checked;
        public int StartMenuSubmenuOpenDelayMs => (int)submenuOpenDelayInput.Value;
        public bool AutoHideTaskbar => autoHideTaskbarCheckBox.Checked;
        public bool StartOnWindowsStartup => startOnWindowsStartupCheckBox.Checked;
        public string TaskbarFontName => fontNameCombo.SelectedItem?.ToString() ?? "\u65B0\u7D30\u660E\u9AD4";
        public float TaskbarFontSize => (float)fontSizeInput.Value;
        public Color TaskbarFontColor => taskbarFontColorRow.Preview.BackColor;
        public string TaskbarButtonStyle => buttonStyleCombo.SelectedItem?.ToString() ?? "Win98";
        public int TaskbarBevelSize => (int)bevelSizeInput.Value;
        public Color TaskbarBaseColor => taskbarBaseColorRow.Preview.BackColor;
        public Color TaskbarLightColor => taskbarLightColorRow.Preview.BackColor;
        public Color TaskbarDarkColor => taskbarDarkColorRow.Preview.BackColor;
        public string SelectedThemeProfileName => themeProfileCombo.SelectedItem?.ToString() ?? "Custom";

        public TaskbarOptionsForm(
            int currentStartMenuIconSize,
            int currentTaskIconSize,
            bool currentLazyLoadProgramsSubmenu,
            bool currentPlayVolumeFeedbackSound,
            bool currentUseClassicVolumePopup,
            int currentStartMenuSubmenuOpenDelayMs,
            bool currentAutoHideTaskbar,
            bool currentStartOnWindowsStartup,
            string currentTaskbarButtonStyle,
            int currentTaskbarBevelSize,
            string currentTaskbarFontName,
            float currentTaskbarFontSize,
            Color currentTaskbarFontColor,
            Color currentTaskbarBaseColor,
            Color currentTaskbarLightColor,
            Color currentTaskbarDarkColor,
            string currentThemeProfileName,
            Action openQuickLaunchFolderAction)
        {
            this.openQuickLaunchFolderAction = openQuickLaunchFolderAction;

            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.Manual;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            TopMost = true;
            Text = "Taskbar Options";
            ClientSize = new Size(640, 520);
            MinimumSize = new Size(640, 520);
            Font = new Font("MS Sans Serif", 8.25f, FontStyle.Regular, GraphicsUnit.Point);

            var desktopArea = Screen.PrimaryScreen?.WorkingArea ?? new Rectangle(0, 0, 1024, 768);
            var x = desktopArea.Left + Math.Max(0, (desktopArea.Width - ClientSize.Width) / 2);
            var y = desktopArea.Top + Math.Max(0, (desktopArea.Height - ClientSize.Height) / 2);
            Location = new Point(x, y);

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                Padding = new Padding(10)
            };
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Controls.Add(root);

            var tabs = new TabControl { Dock = DockStyle.Fill };
            root.Controls.Add(tabs, 0, 0);

            var appearanceTab = new TabPage("Appearance");
            var behaviorTab = new TabPage("Behavior");
            var startMenuTab = new TabPage("Start Menu");
            var trayTab = new TabPage("Tray");
            tabs.TabPages.Add(appearanceTab);
            tabs.TabPages.Add(behaviorTab);
            tabs.TabPages.Add(startMenuTab);
            tabs.TabPages.Add(trayTab);

            var appearanceLayout = CreateTabLayout();
            appearanceTab.Controls.Add(appearanceLayout);

            themeProfileCombo = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Width = 180
            };
            themeProfileCombo.Items.AddRange(new object[] { "Custom", "Win95", "Win98 Classic", "Win98 Plus-like" });
            themeProfileCombo.SelectedItem = themeProfileCombo.Items.Contains(currentThemeProfileName) ? currentThemeProfileName : "Custom";
            var themeButtons = new FlowLayoutPanel
            {
                AutoSize = true,
                WrapContents = true,
                FlowDirection = FlowDirection.LeftToRight,
                Margin = new Padding(0, 4, 0, 0)
            };
            var applyThemeButton = new Button { Text = "Apply Preset", AutoSize = true };
            var resetWin98Button = new Button { Text = "Restore Win98 Defaults", AutoSize = true };
            var resetAllButton = new Button { Text = "Reset All", AutoSize = true };
            var copyThemeButton = new Button { Text = "Copy Theme", AutoSize = true };
            var pasteThemeButton = new Button { Text = "Paste Theme", AutoSize = true };
            applyThemeButton.Click += (_, _) => ApplySelectedThemePreset();
            resetWin98Button.Click += (_, _) => RestoreWin98Defaults();
            resetAllButton.Click += (_, _) => ResetAllToDefaults();
            copyThemeButton.Click += (_, _) => CopyThemeToClipboard();
            pasteThemeButton.Click += (_, _) => PasteThemeFromClipboard();
            themeButtons.Controls.Add(applyThemeButton);
            themeButtons.Controls.Add(resetWin98Button);
            themeButtons.Controls.Add(resetAllButton);
            themeButtons.Controls.Add(copyThemeButton);
            themeButtons.Controls.Add(pasteThemeButton);
            AddLabeledControl(appearanceLayout, "Theme profile:", CreateStackedContainer(themeProfileCombo, themeButtons));

            fontNameCombo = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 240 };
            var fontNames = FontFamily.Families.Select(f => f.Name).OrderBy(n => n, StringComparer.CurrentCultureIgnoreCase).ToList();
            fontNameCombo.Items.AddRange(fontNames.Cast<object>().ToArray());
            if (fontNames.Contains(currentTaskbarFontName, StringComparer.CurrentCultureIgnoreCase))
            {
                fontNameCombo.SelectedItem = fontNames.First(name => name.Equals(currentTaskbarFontName, StringComparison.CurrentCultureIgnoreCase));
            }
            else
            {
                fontNameCombo.Items.Insert(0, currentTaskbarFontName);
                fontNameCombo.SelectedIndex = 0;
            }

            fontSizeInput = new NumericUpDown
            {
                Minimum = 7,
                Maximum = 16,
                DecimalPlaces = 2,
                Increment = 0.25m,
                Value = Math.Clamp((decimal)currentTaskbarFontSize, 7m, 16m),
                Width = 90
            };

            buttonStyleCombo = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 140 };
            buttonStyleCombo.Items.AddRange(new object[] { "Modern", "Win98" });
            buttonStyleCombo.SelectedItem = buttonStyleCombo.Items.Contains(currentTaskbarButtonStyle) ? currentTaskbarButtonStyle : "Win98";

            bevelSizeInput = new NumericUpDown
            {
                Minimum = 1,
                Maximum = 4,
                Value = Math.Clamp(currentTaskbarBevelSize, 1, 4),
                Width = 90
            };

            AddLabeledControl(appearanceLayout, "Taskbar font:", fontNameCombo);
            AddLabeledControl(appearanceLayout, "Taskbar font size:", fontSizeInput);
            AddLabeledControl(appearanceLayout, "Taskbar button style:", buttonStyleCombo);
            AddLabeledControl(appearanceLayout, "Bevel size (px):", bevelSizeInput);

            taskbarFontColorRow = CreateColorRow(currentTaskbarFontColor, Color.Black);
            taskbarBaseColorRow = CreateColorRow(currentTaskbarBaseColor, Color.FromArgb(192, 192, 192));
            taskbarLightColorRow = CreateColorRow(currentTaskbarLightColor, Color.White);
            taskbarDarkColorRow = CreateColorRow(currentTaskbarDarkColor, Color.FromArgb(128, 128, 128));

            AddLabeledControl(appearanceLayout, "Taskbar font color:", ComposeColorRow(taskbarFontColorRow));
            AddLabeledControl(appearanceLayout, "Taskbar color:", ComposeColorRow(taskbarBaseColorRow));
            AddLabeledControl(appearanceLayout, "Light color:", ComposeColorRow(taskbarLightColorRow));
            AddLabeledControl(appearanceLayout, "Dark color:", ComposeColorRow(taskbarDarkColorRow));

            var behaviorLayout = CreateTabLayout();
            behaviorTab.Controls.Add(behaviorLayout);

            startMenuIconSizeInput = new NumericUpDown
            {
                Minimum = 16,
                Maximum = 32,
                Value = Math.Clamp(currentStartMenuIconSize, 16, 32),
                Width = 90
            };
            taskIconSizeInput = new NumericUpDown
            {
                Minimum = 16,
                Maximum = 32,
                Value = Math.Clamp(currentTaskIconSize, 16, 32),
                Width = 90
            };
            autoHideTaskbarCheckBox = new CheckBox { Text = "Auto-hide taskbar", Checked = currentAutoHideTaskbar, AutoSize = true };
            startOnWindowsStartupCheckBox = new CheckBox { Text = "Start on Windows startup", Checked = currentStartOnWindowsStartup, AutoSize = true };
            var exportButton = new Button { Text = "Export Settings...", AutoSize = true };
            var importButton = new Button { Text = "Import Settings...", AutoSize = true };
            exportButton.Click += (_, _) => ExportSettingsToJson();
            importButton.Click += (_, _) => ImportSettingsFromJson();

            AddLabeledControl(behaviorLayout, "Start/menu icon size:", startMenuIconSizeInput);
            AddLabeledControl(behaviorLayout, "Task icon size:", taskIconSizeInput);
            AddLabeledControl(behaviorLayout, "Taskbar visibility:", autoHideTaskbarCheckBox);
            AddLabeledControl(behaviorLayout, "Startup:", startOnWindowsStartupCheckBox);
            AddLabeledControl(behaviorLayout, "Settings exchange:", CreateRowContainer(exportButton, importButton));

            var startMenuLayout = CreateTabLayout();
            startMenuTab.Controls.Add(startMenuLayout);
            submenuOpenDelayInput = new NumericUpDown
            {
                Minimum = 0,
                Maximum = 1500,
                Increment = 25,
                Value = Math.Clamp(currentStartMenuSubmenuOpenDelayMs, 0, 1500),
                Width = 90
            };
            lazyLoadProgramsCheckBox = new CheckBox
            {
                Text = "Lazy-load Programs submenu",
                Checked = currentLazyLoadProgramsSubmenu,
                AutoSize = true
            };
            AddLabeledControl(startMenuLayout, "Submenu open delay (ms):", submenuOpenDelayInput);
            AddLabeledControl(startMenuLayout, "Programs loading:", lazyLoadProgramsCheckBox);
            var openUserProgramsButton = new Button { Text = "Open User Programs Folder...", AutoSize = true };
            var openCommonProgramsButton = new Button { Text = "Open Common Programs Folder...", AutoSize = true };
            var openStartupFolderButton = new Button { Text = "Open Startup Folder...", AutoSize = true };
            openUserProgramsButton.Click += (_, _) => OpenSpecialFolder(Environment.SpecialFolder.Programs);
            openCommonProgramsButton.Click += (_, _) => OpenSpecialFolder(Environment.SpecialFolder.CommonPrograms);
            openStartupFolderButton.Click += (_, _) => OpenSpecialFolder(Environment.SpecialFolder.Startup);
            AddLabeledControl(startMenuLayout, "Start menu folders:", CreateStackedContainer(
                CreateRowContainer(openUserProgramsButton, openCommonProgramsButton),
                openStartupFolderButton));

            var trayLayout = CreateTabLayout();
            trayTab.Controls.Add(trayLayout);
            playVolumeFeedbackSoundCheckBox = new CheckBox
            {
                Text = "Play feedback sound on volume slider release",
                Checked = currentPlayVolumeFeedbackSound,
                AutoSize = true
            };
            useClassicVolumePopupCheckBox = new CheckBox
            {
                Text = "Use classic built-in volume popup (instead of sndvol.exe)",
                Checked = currentUseClassicVolumePopup,
                AutoSize = true
            };
            var openQuickLaunchFolderButton = new Button
            {
                Text = "Open Quick Launch Folder...",
                AutoSize = true
            };
            openQuickLaunchFolderButton.Click += (_, _) => this.openQuickLaunchFolderAction();
            var openNotificationSettingsButton = new Button
            {
                Text = "Open Windows Notification Settings...",
                AutoSize = true
            };
            openNotificationSettingsButton.Click += (_, _) => OpenUriOrFallback("ms-settings:notifications", "control.exe");
            AddLabeledControl(trayLayout, "Volume feedback:", playVolumeFeedbackSoundCheckBox);
            AddLabeledControl(trayLayout, "Volume UI:", useClassicVolumePopupCheckBox);
            AddLabeledControl(trayLayout, "Quick Launch:", openQuickLaunchFolderButton);
            AddLabeledControl(trayLayout, "Notifications:", openNotificationSettingsButton);

            var buttons = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Fill,
                AutoSize = true,
                WrapContents = false
            };
            var btnApply = new Button { Text = "Apply", Width = 88 };
            var btnCancel = new Button { Text = "Cancel", Width = 88, DialogResult = DialogResult.Cancel };
            var btnOk = new Button { Text = "OK", Width = 88, DialogResult = DialogResult.OK };
            btnApply.Click += (_, _) => ApplyRequested?.Invoke();
            buttons.Controls.Add(btnApply);
            buttons.Controls.Add(btnCancel);
            buttons.Controls.Add(btnOk);
            root.Controls.Add(buttons, 0, 1);

            AcceptButton = btnOk;
            CancelButton = btnCancel;
        }

        private static TableLayoutPanel CreateTabLayout()
        {
            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                ColumnCount = 2,
                RowCount = 0,
                GrowStyle = TableLayoutPanelGrowStyle.AddRows,
                Padding = new Padding(8)
            };

            layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            return layout;
        }

        private static void AddLabeledControl(TableLayoutPanel layout, string label, Control control)
        {
            var row = layout.RowCount++;
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var caption = new Label
            {
                Text = label,
                AutoSize = true,
                Anchor = AnchorStyles.Left | AnchorStyles.Top,
                Margin = new Padding(0, 8, 8, 8)
            };

            control.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;
            control.Margin = new Padding(0, 4, 0, 4);

            layout.Controls.Add(caption, 0, row);
            layout.Controls.Add(control, 1, row);
        }

        private static Control CreateRowContainer(params Control[] controls)
        {
            var panel = new FlowLayoutPanel
            {
                AutoSize = true,
                WrapContents = false,
                FlowDirection = FlowDirection.LeftToRight,
                Margin = Padding.Empty
            };

            foreach (var control in controls)
            {
                panel.Controls.Add(control);
            }

            return panel;
        }

        private static Control CreateStackedContainer(Control top, Control bottom)
        {
            var panel = new TableLayoutPanel
            {
                AutoSize = true,
                ColumnCount = 1,
                RowCount = 2,
                GrowStyle = TableLayoutPanelGrowStyle.FixedSize,
                Margin = Padding.Empty,
                Padding = Padding.Empty
            };
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            top.Margin = new Padding(0);
            bottom.Margin = new Padding(0);
            panel.Controls.Add(top, 0, 0);
            panel.Controls.Add(bottom, 0, 1);
            return panel;
        }

        private static void OpenSpecialFolder(Environment.SpecialFolder folder)
        {
            var path = Environment.GetFolderPath(folder);
            if (string.IsNullOrWhiteSpace(path) || !Directory.Exists(path))
            {
                System.Media.SystemSounds.Beep.Play();
                return;
            }

            OpenPath(path);
        }

        private static void OpenUriOrFallback(string uri, string fallbackFileName)
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = uri,
                    UseShellExecute = true
                });
                return;
            }
            catch
            {
            }

            OpenPath(fallbackFileName);
        }

        private static void OpenPath(string path)
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = path,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Unable to open '{path}'.{Environment.NewLine}{ex.Message}", "Taskbar Options", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static string ToHex(Color color) => ColorTranslator.ToHtml(color).ToUpperInvariant();

        private ColorRowControls CreateColorRow(Color value, Color fallback)
        {
            var preview = new Panel
            {
                Width = 48,
                Height = 22,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = value
            };
            var hex = new TextBox { Width = 84, Text = ToHex(value) };
            var pickButton = new Button { Text = "...", Width = 34 };
            var defaultButton = new Button { Text = "Default", Width = 64 };

            pickButton.Click += (_, _) =>
            {
                using var dialog = new ColorDialog { FullOpen = true, Color = preview.BackColor };
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    preview.BackColor = dialog.Color;
                    hex.Text = ToHex(dialog.Color);
                    themeProfileCombo.SelectedItem = "Custom";
                }
            };

            defaultButton.Click += (_, _) =>
            {
                preview.BackColor = fallback;
                hex.Text = ToHex(fallback);
                themeProfileCombo.SelectedItem = "Custom";
            };

            hex.Leave += (_, _) =>
            {
                if (TryParseColor(hex.Text, out var parsed))
                {
                    preview.BackColor = parsed;
                    hex.Text = ToHex(parsed);
                    themeProfileCombo.SelectedItem = "Custom";
                }
                else
                {
                    hex.Text = ToHex(preview.BackColor);
                }
            };

            preview.Tag = CreateRowContainer(preview, hex, pickButton, defaultButton);
            return new ColorRowControls(preview, hex);
        }

        private static Control ComposeColorRow(ColorRowControls row)
        {
            return row.Preview.Tag as Control ?? row.Preview;
        }

        private static bool TryParseColor(string input, out Color color)
        {
            try
            {
                color = ColorTranslator.FromHtml(input.Trim());
                return true;
            }
            catch
            {
                color = Color.Empty;
                return false;
            }
        }

        private void ApplySelectedThemePreset()
        {
            var selected = SelectedThemeProfileName;
            ApplyThemePreset(selected);
        }

        private void RestoreWin98Defaults()
        {
            themeProfileCombo.SelectedItem = "Win98 Classic";
            ApplyThemePreset("Win98 Classic");
        }

        private void ApplyThemePreset(string preset)
        {
            switch (preset)
            {
                case "Win95":
                    buttonStyleCombo.SelectedItem = "Win98";
                    bevelSizeInput.Value = 1;
                    fontNameCombo.SelectedItem = "MS Sans Serif";
                    fontSizeInput.Value = 8.25m;
                    SetColorRow(taskbarFontColorRow, Color.Black);
                    SetColorRow(taskbarBaseColorRow, Color.FromArgb(192, 192, 192));
                    SetColorRow(taskbarLightColorRow, Color.White);
                    SetColorRow(taskbarDarkColorRow, Color.FromArgb(128, 128, 128));
                    break;
                case "Win98 Plus-like":
                    buttonStyleCombo.SelectedItem = "Win98";
                    bevelSizeInput.Value = 2;
                    fontNameCombo.SelectedItem = "MS Sans Serif";
                    fontSizeInput.Value = 8.25m;
                    SetColorRow(taskbarFontColorRow, Color.FromArgb(10, 10, 10));
                    SetColorRow(taskbarBaseColorRow, Color.FromArgb(172, 183, 200));
                    SetColorRow(taskbarLightColorRow, Color.FromArgb(238, 243, 250));
                    SetColorRow(taskbarDarkColorRow, Color.FromArgb(88, 100, 124));
                    break;
                case "Win98 Classic":
                default:
                    buttonStyleCombo.SelectedItem = "Win98";
                    bevelSizeInput.Value = 2;
                    fontNameCombo.SelectedItem = "MS Sans Serif";
                    fontSizeInput.Value = 8.25m;
                    SetColorRow(taskbarFontColorRow, Color.Black);
                    SetColorRow(taskbarBaseColorRow, Color.FromArgb(192, 192, 192));
                    SetColorRow(taskbarLightColorRow, Color.White);
                    SetColorRow(taskbarDarkColorRow, Color.FromArgb(128, 128, 128));
                    break;
            }
        }

        private static void SetColorRow(ColorRowControls row, Color value)
        {
            row.Preview.BackColor = value;
            row.HexBox.Text = ToHex(value);
        }

        private void ResetAllToDefaults()
        {
            var defaults = new TaskbarSettings();

            startMenuIconSizeInput.Value = Math.Clamp(defaults.StartMenuIconSize, 16, 32);
            taskIconSizeInput.Value = Math.Clamp(defaults.TaskIconSize, 16, 32);
            lazyLoadProgramsCheckBox.Checked = defaults.LazyLoadProgramsSubmenu;
            playVolumeFeedbackSoundCheckBox.Checked = defaults.PlayVolumeFeedbackSound;
            useClassicVolumePopupCheckBox.Checked = defaults.UseClassicVolumePopup;
            submenuOpenDelayInput.Value = Math.Clamp(defaults.StartMenuSubmenuOpenDelayMs, 0, 1500);
            autoHideTaskbarCheckBox.Checked = defaults.AutoHideTaskbar;
            startOnWindowsStartupCheckBox.Checked = defaults.StartOnWindowsStartup;

            themeProfileCombo.SelectedItem = themeProfileCombo.Items.Contains(defaults.ThemeProfileName)
                ? defaults.ThemeProfileName
                : "Custom";
            buttonStyleCombo.SelectedItem = buttonStyleCombo.Items.Contains(defaults.TaskbarButtonStyle)
                ? defaults.TaskbarButtonStyle
                : "Modern";
            bevelSizeInput.Value = Math.Clamp(defaults.TaskbarBevelSize, 1, 4);

            if (fontNameCombo.Items.Contains(defaults.TaskbarFontName))
            {
                fontNameCombo.SelectedItem = defaults.TaskbarFontName;
            }

            fontSizeInput.Value = Math.Clamp((decimal)defaults.TaskbarFontSize, 7m, 16m);

            if (TryParseColor(defaults.TaskbarFontColor, out var fontColor))
            {
                SetColorRow(taskbarFontColorRow, fontColor);
            }

            if (TryParseColor(defaults.TaskbarBaseColor, out var baseColor))
            {
                SetColorRow(taskbarBaseColorRow, baseColor);
            }

            if (TryParseColor(defaults.TaskbarLightColor, out var lightColor))
            {
                SetColorRow(taskbarLightColorRow, lightColor);
            }

            if (TryParseColor(defaults.TaskbarDarkColor, out var darkColor))
            {
                SetColorRow(taskbarDarkColorRow, darkColor);
            }
        }

        private void CopyThemeToClipboard()
        {
            var payload = new ThemeClipboardPayload(
                SelectedThemeProfileName,
                TaskbarButtonStyle,
                TaskbarBevelSize,
                TaskbarFontName,
                TaskbarFontSize,
                ToHex(TaskbarFontColor),
                ToHex(TaskbarBaseColor),
                ToHex(TaskbarLightColor),
                ToHex(TaskbarDarkColor));
            Clipboard.SetText(JsonSerializer.Serialize(payload, new JsonSerializerOptions { WriteIndented = true }));
        }

        private void PasteThemeFromClipboard()
        {
            try
            {
                if (!Clipboard.ContainsText())
                {
                    return;
                }

                var json = Clipboard.GetText();
                var payload = JsonSerializer.Deserialize<ThemeClipboardPayload>(json);
                if (payload == null)
                {
                    return;
                }

                themeProfileCombo.SelectedItem = themeProfileCombo.Items.Contains(payload.ThemeProfileName) ? payload.ThemeProfileName : "Custom";
                buttonStyleCombo.SelectedItem = buttonStyleCombo.Items.Contains(payload.TaskbarButtonStyle) ? payload.TaskbarButtonStyle : "Win98";
                bevelSizeInput.Value = Math.Clamp(payload.TaskbarBevelSize, 1, 4);

                if (fontNameCombo.Items.Contains(payload.TaskbarFontName))
                {
                    fontNameCombo.SelectedItem = payload.TaskbarFontName;
                }

                fontSizeInput.Value = Math.Clamp((decimal)payload.TaskbarFontSize, 7m, 16m);
                if (TryParseColor(payload.TaskbarFontColor, out var fontColor)) SetColorRow(taskbarFontColorRow, fontColor);
                if (TryParseColor(payload.TaskbarBaseColor, out var baseColor)) SetColorRow(taskbarBaseColorRow, baseColor);
                if (TryParseColor(payload.TaskbarLightColor, out var lightColor)) SetColorRow(taskbarLightColorRow, lightColor);
                if (TryParseColor(payload.TaskbarDarkColor, out var darkColor)) SetColorRow(taskbarDarkColorRow, darkColor);
            }
            catch
            {
            }
        }

        private void ExportSettingsToJson()
        {
            using var dialog = new SaveFileDialog
            {
                Filter = "JSON files (*.json)|*.json",
                FileName = $"win9xplorer-taskbar-{DateTime.Now:yyyyMMdd-HHmmss}.json"
            };
            if (dialog.ShowDialog(this) != DialogResult.OK)
            {
                return;
            }

            var current = TaskbarSettingsStore.Load();
            current.StartMenuIconSize = StartMenuIconSize;
            current.TaskIconSize = TaskIconSize;
            current.LazyLoadProgramsSubmenu = LazyLoadProgramsSubmenu;
            current.PlayVolumeFeedbackSound = PlayVolumeFeedbackSound;
            current.UseClassicVolumePopup = UseClassicVolumePopup;
            current.StartMenuSubmenuOpenDelayMs = StartMenuSubmenuOpenDelayMs;
            current.AutoHideTaskbar = AutoHideTaskbar;
            current.StartOnWindowsStartup = StartOnWindowsStartup;
            current.TaskbarButtonStyle = TaskbarButtonStyle;
            current.TaskbarFontName = TaskbarFontName;
            current.TaskbarFontSize = TaskbarFontSize;
            current.TaskbarFontColor = ToHex(TaskbarFontColor);
            current.TaskbarBaseColor = ToHex(TaskbarBaseColor);
            current.TaskbarLightColor = ToHex(TaskbarLightColor);
            current.TaskbarDarkColor = ToHex(TaskbarDarkColor);
            current.TaskbarBevelSize = TaskbarBevelSize;
            current.ThemeProfileName = SelectedThemeProfileName;

            File.WriteAllText(dialog.FileName, JsonSerializer.Serialize(current, new JsonSerializerOptions { WriteIndented = true }));
        }

        private void ImportSettingsFromJson()
        {
            using var dialog = new OpenFileDialog
            {
                Filter = "JSON files (*.json)|*.json"
            };
            if (dialog.ShowDialog(this) != DialogResult.OK)
            {
                return;
            }

            try
            {
                var json = File.ReadAllText(dialog.FileName);
                var imported = JsonSerializer.Deserialize<TaskbarSettings>(json);
                if (imported == null)
                {
                    return;
                }

                startMenuIconSizeInput.Value = Math.Clamp(imported.StartMenuIconSize, 16, 32);
                taskIconSizeInput.Value = Math.Clamp(imported.TaskIconSize, 16, 32);
                lazyLoadProgramsCheckBox.Checked = imported.LazyLoadProgramsSubmenu;
                playVolumeFeedbackSoundCheckBox.Checked = imported.PlayVolumeFeedbackSound;
                useClassicVolumePopupCheckBox.Checked = imported.UseClassicVolumePopup;
                submenuOpenDelayInput.Value = Math.Clamp(imported.StartMenuSubmenuOpenDelayMs, 0, 1500);
                autoHideTaskbarCheckBox.Checked = imported.AutoHideTaskbar;
                startOnWindowsStartupCheckBox.Checked = imported.StartOnWindowsStartup;
                buttonStyleCombo.SelectedItem = buttonStyleCombo.Items.Contains(imported.TaskbarButtonStyle) ? imported.TaskbarButtonStyle : "Win98";
                bevelSizeInput.Value = Math.Clamp(imported.TaskbarBevelSize, 1, 4);
                if (fontNameCombo.Items.Contains(imported.TaskbarFontName))
                {
                    fontNameCombo.SelectedItem = imported.TaskbarFontName;
                }

                fontSizeInput.Value = Math.Clamp((decimal)imported.TaskbarFontSize, 7m, 16m);
                if (TryParseColor(imported.TaskbarFontColor, out var fontColor)) SetColorRow(taskbarFontColorRow, fontColor);
                if (TryParseColor(imported.TaskbarBaseColor, out var baseColor)) SetColorRow(taskbarBaseColorRow, baseColor);
                if (TryParseColor(imported.TaskbarLightColor, out var lightColor)) SetColorRow(taskbarLightColorRow, lightColor);
                if (TryParseColor(imported.TaskbarDarkColor, out var darkColor)) SetColorRow(taskbarDarkColorRow, darkColor);
                themeProfileCombo.SelectedItem = themeProfileCombo.Items.Contains(imported.ThemeProfileName) ? imported.ThemeProfileName : "Custom";
            }
            catch
            {
            }
        }
    }
}
