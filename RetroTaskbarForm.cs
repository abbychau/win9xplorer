using System.Diagnostics;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Windows.Media.Imaging;
using ManagedShell;
using ManagedShell.WindowsTray;
using ManagedShell.WindowsTasks;
using MediaImageSource = System.Windows.Media.ImageSource;
using TrayNotifyIcon = ManagedShell.WindowsTray.NotifyIcon;
using WpfMouseButton = System.Windows.Input.MouseButton;

namespace win9xplorer
{
    internal sealed class RetroTaskbarForm : Form
    {
        private static RetroTaskbarForm? instance;
        private const int TaskbarHeight = 30;
        private const int RefreshIntervalMs = 120;
        private const int DefaultStartMenuIconSize = 20;
        private const int DefaultTaskIconSize = 20;
        private const int QuickLaunchButtonSize = TaskbarHeight;
        private readonly ShellManager shellManager;
        private readonly Button btnStart;
        private readonly FlowLayoutPanel quickLaunchPanel;
        private readonly FlowLayoutPanel taskButtonsPanel;
        private readonly FlowLayoutPanel trayIconsPanel;
        private readonly Label lblClock;
        private readonly System.Windows.Forms.Timer refreshTimer;
        private ContextMenuStrip startMenu;
        private readonly ContextMenuStrip taskWindowMenu;
        private readonly ContextMenuStrip taskbarContextMenu;
        private readonly ContextMenuStrip clockContextMenu;
        private readonly ToolTip taskToolTip;
        private readonly ToolTip trayToolTip;
        private readonly ToolTip quickLaunchToolTip;
        private readonly Font normalFont;
        private readonly Font activeFont;
        private readonly Image folderMenuIcon;
        private readonly Image startButtonMenuIcon;
        private readonly Image programsMenuIcon;
        private readonly Image explorerMenuIcon;
        private readonly Image documentsMenuIcon;
        private readonly Image settingsMenuIcon;
        private readonly Image runMenuIcon;
        private readonly Image quickLaunchDesktopIcon;
        private readonly Image customVolumeIconImage;
        private sealed record ProgramMenuNode(string FolderPath, int Depth);
        private readonly Dictionary<IntPtr, Button> taskButtons = new();
        private readonly Dictionary<IntPtr, ApplicationWindow> windowsByHandle = new();
        private readonly Dictionary<TrayNotifyIcon, PictureBox> trayIconControls = new();
        private IntPtr contextMenuWindowHandle = IntPtr.Zero;
        private IntPtr lastActiveWindowHandle = IntPtr.Zero;
        private IntPtr foregroundWindowBeforeTaskClick = IntPtr.Zero;
        private readonly WinApi.LowLevelKeyboardProc keyboardProc;
        private IntPtr keyboardHookHandle = IntPtr.Zero;
        private bool winKeyPressed;
        private bool nonWinKeyPressedWhileWinHeld;
        private readonly VolumePopupForm volumePopup;
        private readonly TimeDetailsForm timeDetailsPopup;
        private readonly PictureBox customVolumeIcon;
        private int startMenuIconSize = DefaultStartMenuIconSize;
        private int taskIconSize = DefaultTaskIconSize;
        private bool lazyLoadProgramsSubmenu = true;

        public RetroTaskbarForm()
        {
            instance = this;
            shellManager = new ShellManager();

            FormBorderStyle = FormBorderStyle.None;
            ShowIcon = false;
            ShowInTaskbar = false;
            TopMost = true;
            StartPosition = FormStartPosition.Manual;
            BackColor = SystemColors.Control;
            normalFont = new Font("MS Sans Serif", 8.25f, FontStyle.Regular, GraphicsUnit.Point);
            activeFont = new Font("MS Sans Serif", 8.25f, FontStyle.Bold, GraphicsUnit.Point);
            Font = normalFont;
            folderMenuIcon = GetSmallShellIcon("folder", isDirectory: true) ?? SystemIcons.WinLogo.ToBitmap();
            startButtonMenuIcon = LoadMenuAsset("Taskbar and Start Menu.png", SystemIcons.WinLogo.ToBitmap());
            programsMenuIcon = LoadMenuAsset("Start Menu Programs.png", SystemIcons.WinLogo.ToBitmap());
            explorerMenuIcon = LoadMenuAsset("Explorer.png", SystemIcons.Application.ToBitmap());
            documentsMenuIcon = LoadMenuAsset("My Documents.png", SystemIcons.Information.ToBitmap());
            settingsMenuIcon = LoadMenuAsset("Additional Settings.png", SystemIcons.Shield.ToBitmap());
            runMenuIcon = LoadMenuAsset("Run.png", SystemIcons.Question.ToBitmap());
            quickLaunchDesktopIcon = LoadMenuAsset("Desktop.png", SystemIcons.Application.ToBitmap());
            customVolumeIconImage = LoadMenuAsset("Audio Devices.png", SystemIcons.Information.ToBitmap());

            btnStart = new Button
            {
                Text = "Start",
                Width = 72,
                Height = 22,
                Dock = DockStyle.Left,
                FlatStyle = FlatStyle.Standard,
                Font = activeFont,
                Margin = Padding.Empty,
                TextAlign = ContentAlignment.MiddleLeft,
                TextImageRelation = TextImageRelation.ImageBeforeText,
                ImageAlign = ContentAlignment.MiddleLeft,
                Image = ResizeIconImage(startButtonMenuIcon, startMenuIconSize)
            };
            btnStart.Click += BtnStart_Click;

            lblClock = new Label
            {
                Dock = DockStyle.Right,
                Width = 72,
                TextAlign = ContentAlignment.MiddleCenter,
                BorderStyle = BorderStyle.None,
                BackColor = SystemColors.Control,
                Margin = Padding.Empty,
                Font = normalFont
            };

            quickLaunchPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Left,
                Width = 0,
                WrapContents = false,
                AutoScroll = false,
                Margin = Padding.Empty,
                Padding = Padding.Empty,
                BackColor = SystemColors.Control,
                BorderStyle = BorderStyle.None
            };

            trayIconsPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Right,
                Width = 96,
                WrapContents = false,
                AutoScroll = false,
                Margin = Padding.Empty,
                Padding = new Padding(2, 6, 2, 6),
                BackColor = SystemColors.Control,
                BorderStyle = BorderStyle.None
            };

            customVolumeIcon = new PictureBox
            {
                Width = 18,
                Height = 18,
                Margin = new Padding(1, 0, 1, 0),
                SizeMode = PictureBoxSizeMode.CenterImage,
                Cursor = Cursors.Default,
                Image = ResizeIconImage(customVolumeIconImage, 16)
            };
            customVolumeIcon.MouseUp += CustomVolumeIcon_MouseUp;

            taskButtonsPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                WrapContents = false,
                AutoScroll = false,
                Margin = Padding.Empty,
                Padding = Padding.Empty,
                BackColor = SystemColors.Control
            };

            startMenu = BuildStartMenu();
            taskWindowMenu = BuildTaskWindowMenu();
            taskbarContextMenu = BuildTaskbarContextMenu();
            clockContextMenu = BuildClockContextMenu();
            taskToolTip = new ToolTip();
            trayToolTip = new ToolTip();
            quickLaunchToolTip = new ToolTip();
            keyboardProc = KeyboardHookCallback;
            volumePopup = new VolumePopupForm();
            timeDetailsPopup = new TimeDetailsForm();
            LoadQuickLaunchButtons();

            Controls.Add(taskButtonsPanel);
            Controls.Add(trayIconsPanel);
            Controls.Add(lblClock);
            Controls.Add(quickLaunchPanel);
            Controls.Add(btnStart);
            trayIconsPanel.Controls.Add(customVolumeIcon);
            trayToolTip.SetToolTip(customVolumeIcon, "Volume");

            MouseUp += TaskbarBackground_MouseUp;
            taskButtonsPanel.MouseUp += TaskButtonsPanel_MouseUp;
            lblClock.MouseUp += LblClock_MouseUp;
            lblClock.MouseClick += LblClock_MouseClick;

            refreshTimer = new System.Windows.Forms.Timer { Interval = RefreshIntervalMs };
            refreshTimer.Tick += (_, _) => RefreshTaskbar();

            Load += RetroTaskbarForm_Load;
            FormClosed += RetroTaskbarForm_FormClosed;
            Paint += RetroTaskbarForm_Paint;
        }

        public static RetroTaskbarForm GetOrCreate()
        {
            if (instance == null || instance.IsDisposed)
            {
                instance = new RetroTaskbarForm();
            }

            return instance;
        }

        private void RetroTaskbarForm_Load(object? sender, EventArgs e)
        {
            TrySetExplorerTaskbarHidden(true);
            InstallKeyboardHook();
            InitializeTrayService();
            SetTaskbarBounds();
            RefreshTaskbar();
            refreshTimer.Start();
        }

        private void RetroTaskbarForm_FormClosed(object? sender, FormClosedEventArgs e)
        {
            instance = null;
            refreshTimer.Stop();
            refreshTimer.Dispose();

            foreach (var button in taskButtons.Values)
            {
                DisposeTaskButton(button, removeFromPanel: false);
            }

            taskButtons.Clear();
            windowsByHandle.Clear();

            foreach (var control in quickLaunchPanel.Controls.Cast<Control>().ToList())
            {
                if (control is Button quickLaunchButton)
                {
                    quickLaunchButton.Click -= QuickLaunchButton_Click;
                    quickLaunchButton.Image?.Dispose();
                }

                control.Dispose();
            }

            btnStart.Image?.Dispose();
            startMenu.Dispose();
            taskWindowMenu.Dispose();
            taskbarContextMenu.Dispose();
            clockContextMenu.Dispose();
            taskToolTip.Dispose();
            trayToolTip.Dispose();
            quickLaunchToolTip.Dispose();
            normalFont.Dispose();
            activeFont.Dispose();
            folderMenuIcon.Dispose();
            startButtonMenuIcon.Dispose();
            programsMenuIcon.Dispose();
            explorerMenuIcon.Dispose();
            documentsMenuIcon.Dispose();
            settingsMenuIcon.Dispose();
            runMenuIcon.Dispose();
            quickLaunchDesktopIcon.Dispose();
            customVolumeIcon.MouseUp -= CustomVolumeIcon_MouseUp;
            customVolumeIcon.Image?.Dispose();
            customVolumeIcon.Dispose();
            customVolumeIconImage.Dispose();
            volumePopup.Dispose();
            timeDetailsPopup.Dispose();
            UninstallKeyboardHook();
            TeardownTrayService();

            TrySetExplorerTaskbarHidden(false);
            shellManager.AppBarManager.SignalGracefulShutdown();
            shellManager.Dispose();
        }

        private void RetroTaskbarForm_Paint(object? sender, PaintEventArgs e)
        {
            using var topDarkPen = new Pen(SystemColors.ControlDark);
            using var topLightPen = new Pen(SystemColors.ControlLightLight);
            e.Graphics.DrawLine(topLightPen, 0, 0, Width, 0);
            e.Graphics.DrawLine(topDarkPen, 0, 1, Width, 1);
        }

        private ContextMenuStrip BuildStartMenu()
        {
            var menu = new ContextMenuStrip
            {
                ShowImageMargin = true,
                Font = normalFont,
                ImageScalingSize = new Size(startMenuIconSize, startMenuIconSize)
            };

            menu.Items.Add(CreateMenuItem("File Explorer", explorerMenuIcon, (_, _) => BringExplorerToFront()));

            var programsItem = CreateMenuItem("Programs", programsMenuIcon, null);
            if (lazyLoadProgramsSubmenu)
            {
                var programsLoadingItem = CreateMenuItem("(Loading...)", null, null);
                programsLoadingItem.Enabled = false;
                programsItem.DropDownItems.Add(programsLoadingItem);
                programsItem.DropDownOpening += (_, _) => PopulateProgramsMenu(programsItem.DropDownItems);
            }
            else
            {
                PopulateProgramsMenu(programsItem.DropDownItems);
            }
            menu.Items.Add(programsItem);

            var documentsItem = CreateMenuItem("Documents", documentsMenuIcon, null);
            var documentsLoadingItem = CreateMenuItem("(Loading...)", null, null);
            documentsLoadingItem.Enabled = false;
            documentsItem.DropDownItems.Add(documentsLoadingItem);
            documentsItem.DropDownOpening += (_, _) => PopulateDocumentsMenu(documentsItem.DropDownItems);
            menu.Items.Add(documentsItem);

            var settingsItem = CreateMenuItem("Settings", settingsMenuIcon, null);
            settingsItem.DropDownItems.Add(CreateMenuItem("Control Panel", SystemIcons.Shield.ToBitmap(), (_, _) => LaunchProcess("control.exe")));
            settingsItem.DropDownItems.Add(CreateMenuItem("Printers", SystemIcons.Application.ToBitmap(), (_, _) => LaunchProcess("control.exe", "printers")));
            settingsItem.DropDownItems.Add(CreateMenuItem("Taskbar", SystemIcons.Application.ToBitmap(), (_, _) => Show()));
            menu.Items.Add(settingsItem);

            menu.Items.Add(CreateMenuItem("Run...", runMenuIcon, (_, _) => LaunchProcess("rundll32.exe", "shell32.dll,#61")));
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add(CreateMenuItem("Hide Taskbar", null, (_, _) => Hide()));
            menu.Items.Add(CreateMenuItem("Shut Down...", SystemIcons.Error.ToBitmap(), (_, _) => ShowShutdownPrompt()));
            menu.Items.Add(CreateMenuItem("Exit win9xplorer", null, (_, _) => Application.Exit()));

            return menu;
        }

        private ToolStripMenuItem CreateMenuItem(string text, Image? image, EventHandler? onClick)
        {
            var item = new ToolStripMenuItem(text)
            {
                Image = image != null ? ResizeIconImage(image, startMenuIconSize) : null
            };

            if (onClick != null)
            {
                item.Click += (_, e) =>
                {
                    startMenu.Close(ToolStripDropDownCloseReason.ItemClicked);
                    onClick(item, e);
                };
            }

            return item;
        }

        private void PopulateProgramsMenu(ToolStripItemCollection items)
        {
            items.Clear();

            var startFolders = new[]
            {
                Environment.GetFolderPath(Environment.SpecialFolder.Programs),
                Environment.GetFolderPath(Environment.SpecialFolder.CommonPrograms)
            };

            var addedAny = false;
            foreach (var folder in startFolders.Distinct(StringComparer.OrdinalIgnoreCase))
            {
                if (!Directory.Exists(folder))
                {
                    continue;
                }

                AddProgramEntries(items, folder, depth: 0);
                addedAny = true;
            }

            if (!addedAny)
            {
                items.Add(CreateMenuItem("(No program shortcuts)", null, null));
            }
        }

        private void AddProgramEntries(ToolStripItemCollection items, string folderPath, int depth)
        {
            if (depth > 2)
            {
                return;
            }

            List<string> directories;
            List<string> files;

            try
            {
                directories = Directory.GetDirectories(folderPath).OrderBy(Path.GetFileName).ToList();
                var executablePatterns = new[] { "*.lnk", "*.exe", "*.bat", "*.cmd", "*.com", "*.pif" };
                files = executablePatterns
                    .SelectMany(pattern => Directory.GetFiles(folderPath, pattern))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(Path.GetFileName)
                    .ToList();
            }
            catch
            {
                return;
            }

            foreach (var directory in directories)
            {
                var subMenu = CreateProgramFolderMenuItem(Path.GetFileName(directory), directory, depth + 1, lazyLoadProgramsSubmenu);
                if (subMenu != null)
                {
                    items.Add(subMenu);
                }
            }

            foreach (var file in files)
            {
                var name = Path.GetFileNameWithoutExtension(file);
                items.Add(CreateMenuItem(name, GetSmallShellIcon(file, isDirectory: false), (_, _) => LaunchProcess(file)));
            }
        }

        private ToolStripMenuItem? CreateProgramFolderMenuItem(string displayName, string folderPath, int depth, bool lazyLoadChildren)
        {
            if (depth > 2 || !Directory.Exists(folderPath))
            {
                return null;
            }

            var item = CreateMenuItem(displayName, folderMenuIcon, null);

            if (lazyLoadChildren)
            {
                item.Tag = new ProgramMenuNode(folderPath, depth);
                var loadingItem = CreateMenuItem("(Loading...)", null, null);
                loadingItem.Enabled = false;
                item.DropDownItems.Add(loadingItem);
                item.DropDownOpening += ProgramFolderItem_DropDownOpening;
            }
            else
            {
                AddProgramEntries(item.DropDownItems, folderPath, depth);
            }

            return item;
        }

        private void ProgramFolderItem_DropDownOpening(object? sender, EventArgs e)
        {
            if (sender is not ToolStripMenuItem item || item.Tag is not ProgramMenuNode node)
            {
                return;
            }

            item.DropDownItems.Clear();
            AddProgramEntries(item.DropDownItems, node.FolderPath, node.Depth);

            if (item.DropDownItems.Count == 0)
            {
                var empty = CreateMenuItem("(Empty)", null, null);
                empty.Enabled = false;
                item.DropDownItems.Add(empty);
            }
        }

        private void PopulateDocumentsMenu(ToolStripItemCollection items)
        {
            items.Clear();

            var myDocuments = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            items.Add(CreateMenuItem("My Documents", SystemIcons.Application.ToBitmap(), (_, _) => OpenFolder(myDocuments)));

            var recent = Environment.GetFolderPath(Environment.SpecialFolder.Recent);
            if (Directory.Exists(recent))
            {
                var recents = Directory.GetFiles(recent, "*.lnk").OrderByDescending(File.GetLastWriteTime).Take(8).ToList();
                if (recents.Count > 0)
                {
                    items.Add(new ToolStripSeparator());
                    foreach (var path in recents)
                    {
                        items.Add(CreateMenuItem(Path.GetFileNameWithoutExtension(path), null, (_, _) => LaunchProcess(path)));
                    }
                }
            }
        }

        private static void OpenFolder(string path)
        {
            if (!string.IsNullOrWhiteSpace(path) && Directory.Exists(path))
            {
                LaunchProcess("explorer.exe", $"\"{path}\"");
            }
        }

        private static void ShowShutdownPrompt()
        {
            var result = MessageBox.Show(
                "Are you sure you want to shut down Windows now?",
                "Shut Down Windows",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button2);

            if (result == DialogResult.Yes)
            {
                LaunchProcess("shutdown.exe", "/s /t 0");
            }
        }

        private static void LaunchProcess(string fileName, string? arguments = null)
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = fileName,
                    Arguments = arguments ?? string.Empty,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to launch {fileName}: {ex.Message}");
            }
        }

        private static Image? GetSmallShellIcon(string pathOrExtension, bool isDirectory)
        {
            var shfi = new WinApi.SHFILEINFO();
            var flags = WinApi.SHGFI_ICON | WinApi.SHGFI_SMALLICON;

            if (isDirectory)
            {
                flags |= WinApi.SHGFI_USEFILEATTRIBUTES;
                WinApi.SHGetFileInfo(pathOrExtension, WinApi.FILE_ATTRIBUTE_DIRECTORY, ref shfi, (uint)Marshal.SizeOf(shfi), flags);
            }
            else
            {
                if (!File.Exists(pathOrExtension))
                {
                    flags |= WinApi.SHGFI_USEFILEATTRIBUTES;
                }

                WinApi.SHGetFileInfo(pathOrExtension, 0, ref shfi, (uint)Marshal.SizeOf(shfi), flags);
            }

            if (shfi.hIcon == IntPtr.Zero)
            {
                return null;
            }

            using var icon = Icon.FromHandle(shfi.hIcon);
            var bitmap = icon.ToBitmap();
            WinApi.DestroyIcon(shfi.hIcon);
            return new Bitmap(bitmap, new Size(16, 16));
        }

        private static Image LoadMenuAsset(string fileName, Image fallback)
        {
            var baseDir = Application.StartupPath;
            var iconPath = Path.Combine(baseDir, "Windows XP Icons", fileName);

            try
            {
                if (File.Exists(iconPath))
                {
                    using var image = Image.FromFile(iconPath);
                    return new Bitmap(image);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to load menu asset '{fileName}': {ex.Message}");
            }

            return new Bitmap(fallback);
        }

        private ContextMenuStrip BuildTaskWindowMenu()
        {
            var menu = new ContextMenuStrip
            {
                ShowImageMargin = false,
                Font = normalFont
            };

            menu.Items.Add("Restore", null, (_, _) => InvokeWindowAction(window => window.Restore()));
            menu.Items.Add("Minimize", null, (_, _) => InvokeWindowAction(window => window.Minimize()));
            menu.Items.Add("Maximize", null, (_, _) => InvokeWindowAction(window => window.Maximize()));
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("Bring To Front", null, (_, _) => InvokeWindowAction(window => window.BringToFront()));
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("Close", null, (_, _) => InvokeWindowAction(window => window.Close()));

            menu.Opening += (_, _) =>
            {
                var hasWindow = windowsByHandle.TryGetValue(contextMenuWindowHandle, out var window);
                menu.Items[0].Enabled = hasWindow && window!.IsMinimized;
                menu.Items[1].Enabled = hasWindow && !window!.IsMinimized;
                menu.Items[2].Enabled = hasWindow;
                menu.Items[4].Enabled = hasWindow;
                menu.Items[6].Enabled = hasWindow;
            };

            return menu;
        }

        private ContextMenuStrip BuildTaskbarContextMenu()
        {
            var menu = new ContextMenuStrip
            {
                ShowImageMargin = false,
                Font = normalFont
            };

            menu.Items.Add("Options...", null, (_, _) => ShowOptionsDialog());
            return menu;
        }

        private ContextMenuStrip BuildClockContextMenu()
        {
            var menu = new ContextMenuStrip
            {
                ShowImageMargin = false,
                Font = normalFont
            };

            menu.Items.Add("Open Clock...", null, (_, _) => ShowTimeDetailsModal());
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("Adjust Date/Time...", null, (_, _) => OpenDateTimeSettings());
            return menu;
        }

        private void LoadQuickLaunchButtons()
        {
            foreach (var control in quickLaunchPanel.Controls.Cast<Control>().ToList())
            {
                if (control is Button quickLaunchButton)
                {
                    quickLaunchButton.Click -= QuickLaunchButton_Click;
                    quickLaunchButton.Image?.Dispose();
                }

                control.Dispose();
            }

            quickLaunchPanel.Controls.Clear();

            var quickLaunchFolder = GetQuickLaunchFolderPath();
            if (Directory.Exists(quickLaunchFolder))
            {
                var quickLaunchFiles = new[] { "*.lnk", "*.url", "*.scf" }
                    .SelectMany(pattern => Directory.GetFiles(quickLaunchFolder, pattern))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(Path.GetFileNameWithoutExtension)
                    .ToList();

                foreach (var shortcutPath in quickLaunchFiles)
                {
                    quickLaunchPanel.Controls.Add(CreateQuickLaunchButton(shortcutPath));
                }
            }

            if (quickLaunchPanel.Controls.Count == 0)
            {
                quickLaunchPanel.Controls.Add(CreateFallbackQuickLaunchButton(
                    text: "File Explorer",
                    icon: explorerMenuIcon,
                    onClick: (_, _) => BringExplorerToFront()));

                quickLaunchPanel.Controls.Add(CreateFallbackQuickLaunchButton(
                    text: "Show Desktop",
                    icon: quickLaunchDesktopIcon,
                    onClick: (_, _) => ShowDesktop()));
            }

            ResizeQuickLaunchPanel();
        }

        private static string GetQuickLaunchFolderPath()
        {
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "Microsoft",
                "Internet Explorer",
                "Quick Launch");
        }

        private Button CreateQuickLaunchButton(string shortcutPath)
        {
            using var icon = GetSmallShellIcon(shortcutPath, isDirectory: false);
            var image = icon != null ? ResizeIconImage(icon, 16) : ResizeIconImage(explorerMenuIcon, 16);
            var displayName = Path.GetFileNameWithoutExtension(shortcutPath);

            var button = new Button
            {
                Width = QuickLaunchButtonSize,
                Height = QuickLaunchButtonSize,
                Margin = new Padding(0),
                Padding = Padding.Empty,
                FlatStyle = FlatStyle.Standard,
                ImageAlign = ContentAlignment.MiddleCenter,
                Image = image,
                Tag = shortcutPath
            };

            quickLaunchToolTip.SetToolTip(button, displayName);
            button.Click += QuickLaunchButton_Click;
            return button;
        }

        private Button CreateFallbackQuickLaunchButton(string text, Image icon, EventHandler onClick)
        {
            var button = new Button
            {
                Width = QuickLaunchButtonSize,
                Height = QuickLaunchButtonSize,
                Margin = new Padding(0),
                Padding = Padding.Empty,
                FlatStyle = FlatStyle.Standard,
                ImageAlign = ContentAlignment.MiddleCenter,
                Image = ResizeIconImage(icon, 16)
            };

            quickLaunchToolTip.SetToolTip(button, text);
            button.Click += onClick;
            return button;
        }

        private void ResizeQuickLaunchPanel()
        {
            var count = quickLaunchPanel.Controls.Count;
            quickLaunchPanel.Width = count > 0 ? (count * QuickLaunchButtonSize) + quickLaunchPanel.Padding.Horizontal : 0;
        }

        private void QuickLaunchButton_Click(object? sender, EventArgs e)
        {
            if (sender is not Button button || button.Tag is not string shortcutPath)
            {
                return;
            }

            LaunchProcess(shortcutPath);
        }

        private void ShowDesktop()
        {
            foreach (var window in windowsByHandle.Values)
            {
                if (!window.IsMinimized)
                {
                    window.Minimize();
                }
            }
        }

        private void LblClock_MouseClick(object? sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left)
            {
                return;
            }

            ShowTimeDetailsModal();
        }

        private void LblClock_MouseUp(object? sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Right)
            {
                return;
            }

            clockContextMenu.Show(lblClock, e.Location);
        }

        private void ShowTimeDetailsModal()
        {
            var clockBounds = lblClock.RectangleToScreen(lblClock.ClientRectangle);
            var popupX = clockBounds.Right - timeDetailsPopup.Width;
            var popupY = clockBounds.Top - timeDetailsPopup.Height - 4;
            var screenBounds = Screen.FromPoint(clockBounds.Location).WorkingArea;

            if (popupX < screenBounds.Left)
            {
                popupX = screenBounds.Left;
            }

            if (popupX + timeDetailsPopup.Width > screenBounds.Right)
            {
                popupX = screenBounds.Right - timeDetailsPopup.Width;
            }

            if (popupY < screenBounds.Top)
            {
                popupY = clockBounds.Bottom + 4;
            }

            if (timeDetailsPopup.Visible)
            {
                timeDetailsPopup.Hide();
                return;
            }

            timeDetailsPopup.ShowAt(new Point(popupX, popupY));
        }

        private static void OpenDateTimeSettings()
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "ms-settings:dateandtime",
                    UseShellExecute = true
                });
                return;
            }
            catch
            {
            }

            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "control.exe",
                    Arguments = "timedate.cpl",
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to open Date/Time settings: {ex.Message}");
            }
        }

        private void TaskbarBackground_MouseUp(object? sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Right)
            {
                return;
            }

            if (GetChildAtPoint(e.Location) != null)
            {
                return;
            }

            taskbarContextMenu.Show(this, e.Location);
        }

        private void TaskButtonsPanel_MouseUp(object? sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Right)
            {
                return;
            }

            if (taskButtonsPanel.GetChildAtPoint(e.Location) != null)
            {
                return;
            }

            taskbarContextMenu.Show(taskButtonsPanel, e.Location);
        }

        private void ShowOptionsDialog()
        {
            using var dialog = new TaskbarOptionsForm(
                startMenuIconSize,
                taskIconSize,
                lazyLoadProgramsSubmenu,
                OpenQuickLaunchFolderInExplorerForm);
            if (dialog.ShowDialog(this) != DialogResult.OK)
            {
                return;
            }

            startMenuIconSize = dialog.StartMenuIconSize;
            taskIconSize = dialog.TaskIconSize;
            lazyLoadProgramsSubmenu = dialog.LazyLoadProgramsSubmenu;
            ApplyIconSizeSettings();
        }

        private void OpenQuickLaunchFolderInExplorerForm()
        {
            var quickLaunchFolder = GetQuickLaunchFolderPath();
            Directory.CreateDirectory(quickLaunchFolder);

            var explorerForm = Application.OpenForms.OfType<Form1>().FirstOrDefault(form => !form.IsDisposed);
            if (explorerForm == null)
            {
                explorerForm = new Form1();
                explorerForm.Show();
            }

            if (explorerForm.WindowState == FormWindowState.Minimized)
            {
                explorerForm.WindowState = FormWindowState.Normal;
            }

            explorerForm.Show();
            explorerForm.BringToFront();
            explorerForm.Activate();
            explorerForm.OpenPath(quickLaunchFolder);
        }

        private void ApplyIconSizeSettings()
        {
            var oldStartImage = btnStart.Image;
            btnStart.Image = ResizeIconImage(startButtonMenuIcon, startMenuIconSize);
            oldStartImage?.Dispose();

            startMenu.Dispose();
            startMenu = BuildStartMenu();

            foreach (var button in taskButtons.Values)
            {
                var oldImage = button.Image;
                button.Image = null;
                oldImage?.Dispose();
            }

            RefreshTaskbar();
        }

        private void InvokeWindowAction(Action<ApplicationWindow> action)
        {
            if (!windowsByHandle.TryGetValue(contextMenuWindowHandle, out var window))
            {
                return;
            }

            action(window);
            RefreshTaskbar();
        }

        private void BtnStart_Click(object? sender, EventArgs e)
        {
            ToggleStartMenu();
        }

        private void ToggleStartMenu()
        {
            if (startMenu.Visible)
            {
                startMenu.Hide();
                return;
            }

            var startButtonLocation = btnStart.PointToScreen(Point.Empty);
            var menuX = startButtonLocation.X;
            var menuY = startButtonLocation.Y - startMenu.Height;
            var screenBounds = Screen.FromPoint(startButtonLocation).WorkingArea;

            if (menuX + startMenu.Width > screenBounds.Right)
            {
                menuX = screenBounds.Right - startMenu.Width;
            }

            if (menuY < screenBounds.Top)
            {
                menuY = screenBounds.Top;
            }

            startMenu.Show(new Point(menuX, menuY));
        }

        private void InstallKeyboardHook()
        {
            if (keyboardHookHandle != IntPtr.Zero)
            {
                return;
            }

            keyboardHookHandle = WinApi.SetWindowsHookEx(WinApi.WH_KEYBOARD_LL, keyboardProc, IntPtr.Zero, 0);
            if (keyboardHookHandle == IntPtr.Zero)
            {
                Debug.WriteLine("Failed to install keyboard hook.");
            }
        }

        private void UninstallKeyboardHook()
        {
            if (keyboardHookHandle == IntPtr.Zero)
            {
                return;
            }

            WinApi.UnhookWindowsHookEx(keyboardHookHandle);
            keyboardHookHandle = IntPtr.Zero;
        }

        private void InitializeTrayService()
        {
            try
            {
                shellManager.NotificationArea.Initialize();
                shellManager.NotificationArea.TrayIcons.CollectionChanged += TrayIcons_CollectionChanged;

                foreach (var icon in shellManager.NotificationArea.TrayIcons)
                {
                    icon.PropertyChanged += TrayIcon_PropertyChanged;
                }

                SyncTrayIcons();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to initialize tray service: {ex.Message}");
            }
        }

        private void TeardownTrayService()
        {
            try
            {
                shellManager.NotificationArea.TrayIcons.CollectionChanged -= TrayIcons_CollectionChanged;

                foreach (var icon in shellManager.NotificationArea.TrayIcons)
                {
                    icon.PropertyChanged -= TrayIcon_PropertyChanged;
                }

                foreach (var control in trayIconControls.Values.ToList())
                {
                    DisposeTrayIconControl(control);
                }

                trayIconControls.Clear();
                trayIconsPanel.Controls.Clear();
                ResizeTrayPanel();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to teardown tray service: {ex.Message}");
            }
        }

        private void TrayIcons_CollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
        {
            if (e.NewItems != null)
            {
                foreach (var item in e.NewItems)
                {
                    if (item is TrayNotifyIcon icon)
                    {
                        icon.PropertyChanged += TrayIcon_PropertyChanged;
                    }
                }
            }

            if (e.OldItems != null)
            {
                foreach (var item in e.OldItems)
                {
                    if (item is TrayNotifyIcon icon)
                    {
                        icon.PropertyChanged -= TrayIcon_PropertyChanged;
                    }
                }
            }

            if (InvokeRequired)
            {
                BeginInvoke(new Action(SyncTrayIcons));
                return;
            }

            SyncTrayIcons();
        }

        private void TrayIcon_PropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (InvokeRequired)
            {
                BeginInvoke(new Action(SyncTrayIcons));
                return;
            }

            SyncTrayIcons();
        }

        private void SyncTrayIcons()
        {
            var currentIcons = shellManager.NotificationArea.TrayIcons.ToList();
            var visibleIcons = currentIcons.Where(icon => !icon.IsHidden && icon.Icon != null && !IsVolumeTrayIcon(icon)).ToList();

            var removedIcons = trayIconControls.Keys.Where(icon => !visibleIcons.Contains(icon)).ToList();
            foreach (var icon in removedIcons)
            {
                if (!trayIconControls.TryGetValue(icon, out var control))
                {
                    continue;
                }

                trayIconsPanel.Controls.Remove(control);
                DisposeTrayIconControl(control);
                trayIconControls.Remove(icon);
            }

            for (var i = 0; i < visibleIcons.Count; i++)
            {
                var icon = visibleIcons[i];
                if (!trayIconControls.TryGetValue(icon, out var control))
                {
                    control = CreateTrayIconControl(icon);
                    trayIconControls[icon] = control;
                    trayIconsPanel.Controls.Add(control);
                }

                control.Tag = icon;
                trayToolTip.SetToolTip(control, string.IsNullOrWhiteSpace(icon.Title) ? icon.Path : icon.Title);
                UpdateTrayIconImage(control, icon);
                UpdateTrayIconPlacement(icon, control);

                if (trayIconsPanel.Controls.GetChildIndex(control) != i)
                {
                    trayIconsPanel.Controls.SetChildIndex(control, i);
                }
            }

            if (!trayIconsPanel.Controls.Contains(customVolumeIcon))
            {
                trayIconsPanel.Controls.Add(customVolumeIcon);
            }

            trayIconsPanel.Controls.SetChildIndex(customVolumeIcon, trayIconsPanel.Controls.Count - 1);

            ResizeTrayPanel();
            UpdateTrayHostSizeData();
        }

        private PictureBox CreateTrayIconControl(TrayNotifyIcon icon)
        {
            var control = new PictureBox
            {
                Width = 18,
                Height = 18,
                Margin = new Padding(1, 0, 1, 0),
                SizeMode = PictureBoxSizeMode.CenterImage,
                Tag = icon,
                Cursor = Cursors.Default
            };

            control.MouseDown += TrayIcon_MouseDown;
            control.MouseUp += TrayIcon_MouseUp;
            control.MouseMove += TrayIcon_MouseMove;
            control.MouseEnter += TrayIcon_MouseEnter;
            control.MouseLeave += TrayIcon_MouseLeave;

            UpdateTrayIconImage(control, icon);
            return control;
        }

        private void UpdateTrayIconImage(PictureBox control, TrayNotifyIcon icon)
        {
            var oldImage = control.Image;
            var image = ConvertImageSourceToBitmap(icon.Icon, 16);
            control.Image = image;
            oldImage?.Dispose();
        }

        private void ResizeTrayPanel()
        {
            var count = trayIconControls.Count + 1;
            trayIconsPanel.Width = Math.Clamp((count * 20) + 8, 32, 220);
        }

        private void UpdateTrayHostSizeData()
        {
            try
            {
                var bounds = trayIconsPanel.RectangleToScreen(trayIconsPanel.ClientRectangle);
                var data = new TrayHostSizeData
                {
                    rc = new ManagedShell.Interop.NativeMethods.Rect(bounds.Left, bounds.Top, bounds.Right, bounds.Bottom),
                    edge = ManagedShell.Interop.NativeMethods.ABEdge.ABE_BOTTOM
                };

                shellManager.NotificationArea.SetTrayHostSizeData(data);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to update tray host size: {ex.Message}");
            }
        }

        private void DisposeTrayIconControl(PictureBox control)
        {
            control.MouseDown -= TrayIcon_MouseDown;
            control.MouseUp -= TrayIcon_MouseUp;
            control.MouseMove -= TrayIcon_MouseMove;
            control.MouseEnter -= TrayIcon_MouseEnter;
            control.MouseLeave -= TrayIcon_MouseLeave;
            trayToolTip.SetToolTip(control, string.Empty);

            var image = control.Image;
            control.Image = null;
            control.Dispose();
            image?.Dispose();
        }

        private void TrayIcon_MouseDown(object? sender, MouseEventArgs e)
        {
            if (sender is not PictureBox control || control.Tag is not TrayNotifyIcon icon)
            {
                return;
            }

            UpdateTrayIconPlacement(icon, control);
            icon.IconMouseDown(ConvertMouseButton(e.Button), GetCursorPositionParam(), SystemInformation.DoubleClickTime);
        }

        private void TrayIcon_MouseUp(object? sender, MouseEventArgs e)
        {
            if (sender is not PictureBox control || control.Tag is not TrayNotifyIcon icon)
            {
                return;
            }

            if (e.Button == MouseButtons.Left && IsVolumeTrayIcon(icon))
            {
                ShowVolumePopupNear(control);
                return;
            }

            UpdateTrayIconPlacement(icon, control);
            icon.IconMouseUp(ConvertMouseButton(e.Button), GetCursorPositionParam(), SystemInformation.DoubleClickTime);
        }

        private void CustomVolumeIcon_MouseUp(object? sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left && e.Button != MouseButtons.Right)
            {
                return;
            }

            ShowVolumePopupNear(customVolumeIcon);
        }

        private void TrayIcon_MouseMove(object? sender, MouseEventArgs e)
        {
            if (sender is not PictureBox control || control.Tag is not TrayNotifyIcon icon)
            {
                return;
            }

            icon.IconMouseMove(GetCursorPositionParam());
        }

        private void TrayIcon_MouseEnter(object? sender, EventArgs e)
        {
            if (sender is not PictureBox control || control.Tag is not TrayNotifyIcon icon)
            {
                return;
            }

            UpdateTrayIconPlacement(icon, control);
            icon.IconMouseEnter(GetCursorPositionParam());
        }

        private void TrayIcon_MouseLeave(object? sender, EventArgs e)
        {
            if (sender is not PictureBox control || control.Tag is not TrayNotifyIcon icon)
            {
                return;
            }

            icon.IconMouseLeave(GetCursorPositionParam());
        }

        private static WpfMouseButton ConvertMouseButton(MouseButtons button)
        {
            return button switch
            {
                MouseButtons.Right => WpfMouseButton.Right,
                MouseButtons.Middle => WpfMouseButton.Middle,
                _ => WpfMouseButton.Left
            };
        }

        private static uint GetCursorPositionParam()
        {
            var cursor = Cursor.Position;
            return (uint)((cursor.X & 0xFFFF) | ((cursor.Y & 0xFFFF) << 16));
        }

        private bool IsVolumeTrayIcon(TrayNotifyIcon icon)
        {
            if (icon.GUID != Guid.Empty &&
                icon.GUID.ToString().Equals(NotificationArea.VOLUME_GUID, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (!string.IsNullOrWhiteSpace(icon.Identifier) &&
                icon.Identifier.Contains("volume", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (!string.IsNullOrWhiteSpace(icon.Path) &&
                (icon.Path.Contains("SndVol", StringComparison.OrdinalIgnoreCase) ||
                 icon.Path.Contains("Audio", StringComparison.OrdinalIgnoreCase) ||
                 icon.Path.Contains("AudioSrv", StringComparison.OrdinalIgnoreCase) ||
                 icon.Path.Contains("SystemSettings", StringComparison.OrdinalIgnoreCase)))
            {
                return true;
            }

            if (!string.IsNullOrWhiteSpace(icon.Title) &&
                icon.Title.Contains("volume", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return false;
        }

        private void ShowVolumePopupNear(Control trayIconControl)
        {
            var iconBounds = trayIconControl.RectangleToScreen(trayIconControl.ClientRectangle);
            var popupX = iconBounds.Left - (volumePopup.Width / 2) + (iconBounds.Width / 2);
            var popupY = iconBounds.Top - volumePopup.Height - 4;
            var screenBounds = Screen.FromPoint(iconBounds.Location).WorkingArea;

            if (popupX < screenBounds.Left)
            {
                popupX = screenBounds.Left;
            }

            if (popupX + volumePopup.Width > screenBounds.Right)
            {
                popupX = screenBounds.Right - volumePopup.Width;
            }

            if (popupY < screenBounds.Top)
            {
                popupY = iconBounds.Bottom + 4;
            }

            if (volumePopup.Visible)
            {
                volumePopup.Hide();
                return;
            }

            volumePopup.ShowAt(new Point(popupX, popupY));
        }

        private static void UpdateTrayIconPlacement(TrayNotifyIcon icon, Control control)
        {
            try
            {
                var bounds = control.RectangleToScreen(control.ClientRectangle);
                icon.Placement = new ManagedShell.Interop.NativeMethods.Rect(bounds.Left, bounds.Top, bounds.Width, bounds.Height);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to update tray icon placement: {ex.Message}");
            }
        }

        private IntPtr KeyboardHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                var keyInfo = Marshal.PtrToStructure<WinApi.KBDLLHOOKSTRUCT>(lParam);
                var vkCode = (int)keyInfo.vkCode;
                var isWinKey = vkCode == WinApi.VK_LWIN || vkCode == WinApi.VK_RWIN;
                var message = wParam.ToInt32();

                if (message == WinApi.WM_KEYDOWN || message == WinApi.WM_SYSKEYDOWN)
                {
                    if (isWinKey)
                    {
                        winKeyPressed = true;
                        nonWinKeyPressedWhileWinHeld = false;
                        return (IntPtr)1;
                    }
                    else if (winKeyPressed)
                    {
                        nonWinKeyPressedWhileWinHeld = true;
                    }
                }
                else if ((message == WinApi.WM_KEYUP || message == WinApi.WM_SYSKEYUP) && isWinKey)
                {
                    var shouldToggleStart = winKeyPressed && !nonWinKeyPressedWhileWinHeld;
                    winKeyPressed = false;
                    nonWinKeyPressedWhileWinHeld = false;

                    if (shouldToggleStart && !IsDisposed)
                    {
                        BeginInvoke(new Action(ToggleStartMenu));
                    }

                    return (IntPtr)1;
                }
            }

            return WinApi.CallNextHookEx(keyboardHookHandle, nCode, wParam, lParam);
        }

        private void BringExplorerToFront()
        {
            var explorer = new Form1();
            explorer.Show();
            explorer.BringToFront();
            explorer.Activate();
        }

        private void RefreshTaskbar()
        {
            SetTaskbarBounds();
            lblClock.Text = DateTime.Now.ToString("h:mm tt");

            var windows = GetVisibleWindows();
            var currentHandles = new HashSet<IntPtr>();
            var foregroundHandle = WinApi.GetForegroundWindow();

            foreach (var window in windows)
            {
                currentHandles.Add(window.Handle);
                windowsByHandle[window.Handle] = window;

                if (!taskButtons.TryGetValue(window.Handle, out var button))
                {
                    button = CreateTaskButton(window.Handle);
                    taskButtons[window.Handle] = button;
                    taskButtonsPanel.Controls.Add(button);
                }

            }

            if (currentHandles.Contains(foregroundHandle))
            {
                lastActiveWindowHandle = foregroundHandle;
            }
            else if (!currentHandles.Contains(lastActiveWindowHandle))
            {
                lastActiveWindowHandle = IntPtr.Zero;
            }

            foreach (var window in windows)
            {
                if (!taskButtons.TryGetValue(window.Handle, out var button))
                {
                    continue;
                }

                UpdateTaskButton(button, window);
            }

            var handlesToRemove = taskButtons.Keys.Where(handle => !currentHandles.Contains(handle)).ToList();
            foreach (var handle in handlesToRemove)
            {
                if (taskButtons.TryGetValue(handle, out var button))
                {
                    DisposeTaskButton(button, removeFromPanel: true);
                    taskButtons.Remove(handle);
                }

                windowsByHandle.Remove(handle);
            }

            ResizeTaskButtons();
        }

        private void DisposeTaskButton(Button button, bool removeFromPanel)
        {
            button.Click -= TaskButton_Click;
            button.MouseDown -= TaskButton_MouseDown;
            button.MouseUp -= TaskButton_MouseUp;

            if (removeFromPanel && taskButtonsPanel.Controls.Contains(button))
            {
                taskButtonsPanel.Controls.Remove(button);
            }

            var image = button.Image;
            button.Image = null;
            button.Dispose();
            image?.Dispose();
        }

        private List<ApplicationWindow> GetVisibleWindows()
        {
            var visibleWindows = new List<ApplicationWindow>();

            var sourceCollection = shellManager.Tasks?.GroupedWindows?.SourceCollection;
            if (sourceCollection is not System.Collections.IEnumerable enumerable)
            {
                return visibleWindows;
            }

            foreach (var item in enumerable)
            {
                if (item is not ApplicationWindow window)
                {
                    continue;
                }

                if (!window.ShowInTaskbar)
                {
                    continue;
                }

                if (string.IsNullOrWhiteSpace(window.Title))
                {
                    continue;
                }

                if (window.Handle == Handle)
                {
                    continue;
                }

                visibleWindows.Add(window);
            }

            return visibleWindows;
        }

        private Button CreateTaskButton(IntPtr handle)
        {
            var button = new Button
            {
                Tag = handle,
                Height = btnStart.Height,
                Width = 160,
                MinimumSize = new Size(btnStart.Width, btnStart.Height),
                Margin = Padding.Empty,
                FlatStyle = FlatStyle.Standard,
                TextAlign = ContentAlignment.MiddleLeft,
                TextImageRelation = TextImageRelation.ImageBeforeText,
                ImageAlign = ContentAlignment.MiddleLeft,
                Font = normalFont
            };

            button.Click += TaskButton_Click;
            button.MouseDown += TaskButton_MouseDown;
            button.MouseUp += TaskButton_MouseUp;
            return button;
        }

        private void UpdateTaskButton(Button button, ApplicationWindow window)
        {
            var title = string.IsNullOrWhiteSpace(window.Title) ? "Untitled" : window.Title.Trim();
            taskToolTip.SetToolTip(button, title);

            if (button.Image == null)
            {
                var image = ConvertImageSourceToBitmap(window.Icon, taskIconSize);
                if (image != null)
                {
                    button.Image = image;
                }
            }

            var isActive = window.Handle == lastActiveWindowHandle || window.State == ApplicationWindow.WindowState.Active;

            if (isActive)
            {
                button.Font = activeFont;
                button.BackColor = SystemColors.ControlLightLight;
            }
            else
            {
                button.Font = normalFont;
                button.BackColor = SystemColors.Control;
            }

            if (window.State == ApplicationWindow.WindowState.Flashing)
            {
                button.BackColor = Color.FromArgb(255, 255, 192);
            }

            button.Text = TruncateTaskButtonText(button, title, button.Font);
        }

        private string TruncateTaskButtonText(Button button, string text, Font font)
        {
            var horizontalPadding = 12;
            var imageSpace = button.Image != null ? taskIconSize + 8 : 0;
            var maxWidth = Math.Max(24, button.ClientSize.Width - horizontalPadding - imageSpace);

            if (TextRenderer.MeasureText(text, font).Width <= maxWidth)
            {
                return text;
            }

            const string ellipsis = "...";
            var left = 0;
            var right = text.Length;

            while (left < right)
            {
                var mid = (left + right + 1) / 2;
                var candidate = text[..mid] + ellipsis;
                if (TextRenderer.MeasureText(candidate, font).Width <= maxWidth)
                {
                    left = mid;
                }
                else
                {
                    right = mid - 1;
                }
            }

            if (left <= 0)
            {
                return ellipsis;
            }

            return text[..left] + ellipsis;
        }

        private static Bitmap? ConvertImageSourceToBitmap(MediaImageSource? source, int iconSize)
        {
            if (source is not BitmapSource bitmapSource)
            {
                return null;
            }

            try
            {
                using var memoryStream = new MemoryStream();
                var encoder = new PngBitmapEncoder();
                encoder.Frames.Add(BitmapFrame.Create(bitmapSource));
                encoder.Save(memoryStream);
                memoryStream.Position = 0;

                using var image = new Bitmap(memoryStream);
                return new Bitmap(image, new Size(iconSize, iconSize));
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to convert task icon: {ex.Message}");
                return null;
            }
        }

        private static Bitmap ResizeIconImage(Image source, int size)
        {
            return new Bitmap(source, new Size(size, size));
        }

        private void ResizeTaskButtons()
        {
            var count = Math.Max(taskButtonsPanel.Controls.Count, 1);
            var availableWidth = Math.Max(taskButtonsPanel.ClientSize.Width - 4, btnStart.Width);
            var widthPerButton = Math.Clamp(availableWidth / count, btnStart.Width, 220);

            foreach (Control control in taskButtonsPanel.Controls)
            {
                control.Width = widthPerButton;
            }
        }

        private void TaskButton_Click(object? sender, EventArgs e)
        {
            if (sender is not Button button || button.Tag is not IntPtr handle)
            {
                return;
            }

            if (!windowsByHandle.TryGetValue(handle, out var window))
            {
                return;
            }

            if (window.IsMinimized)
            {
                window.Restore();
                window.BringToFront();
            }
            else if (handle == foregroundWindowBeforeTaskClick || handle == lastActiveWindowHandle)
            {
                window.Minimize();
            }
            else
            {
                window.BringToFront();
            }

            foregroundWindowBeforeTaskClick = IntPtr.Zero;

            RefreshTaskbar();
        }

        private void TaskButton_MouseDown(object? sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left)
            {
                return;
            }

            foregroundWindowBeforeTaskClick = WinApi.GetForegroundWindow();
        }

        private void TaskButton_MouseUp(object? sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Right || sender is not Button button || button.Tag is not IntPtr handle)
            {
                return;
            }

            contextMenuWindowHandle = handle;
            taskWindowMenu.Show(button, e.Location);
        }

        private void SetTaskbarBounds()
        {
            var screenBounds = Screen.PrimaryScreen?.Bounds ?? new Rectangle(0, 0, 1024, 768);
            var newBounds = new Rectangle(screenBounds.Left, screenBounds.Bottom - TaskbarHeight, screenBounds.Width, TaskbarHeight);

            if (Bounds != newBounds)
            {
                Bounds = newBounds;
                UpdateTrayHostSizeData();
            }
        }

        private void TrySetExplorerTaskbarHidden(bool hidden)
        {
            try
            {
                shellManager.ExplorerHelper.HideExplorerTaskbar = hidden;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Unable to set Explorer taskbar hidden={hidden}: {ex.Message}");
            }
        }
    }
}
