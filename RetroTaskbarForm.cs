using System.Diagnostics;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;
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
        private static readonly ToolStripRenderer Win9xMenuRenderer = new Win9xMenuRendererImpl();
        private const int TaskbarRowHeight = 30;
        private const int TaskbarButtonVerticalPadding = 2;
        private const int TaskbarBottomPadding = 0;
        private const int TaskbarButtonGap = 2;
        private const int ResizeGripMinHeight = 1;
        private const int RefreshIntervalMs = 120;
        private const int DefaultStartMenuIconSize = 16;
        private const int DefaultTaskIconSize = 16;
        private const int ProgramMenuMaxDepth = 4;
        private const int StartMenuSearchResetSeconds = 2;
        private const int MultiColumnThreshold = 22;
        private const int MaxSubmenuColumns = 3;
        private static readonly string[] ProgramLaunchPatterns =
        {
            "*.lnk",
            "*.exe",
            "*.bat",
            "*.cmd",
            "*.com",
            "*.pif",
            "*.url",
            "*.appref-ms",
            "*.msc",
            "*.cpl"
        };
        private readonly ShellManager shellManager;
        private readonly Panel resizeGripPanel;
        private readonly Panel contentPanel;
        private readonly Panel startHostPanel;
        private readonly Panel notificationAreaPanel;
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
        private readonly ContextMenuStrip quickLaunchItemContextMenu;
        private readonly ToolTip taskToolTip;
        private readonly ToolTip trayToolTip;
        private readonly ToolTip quickLaunchToolTip;
        private Font normalFont;
        private Font activeFont;
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
        private bool startMenuWasVisibleOnStartMouseDown;
        private DateTime lastStartMenuClosedAtUtc = DateTime.MinValue;
        private IntPtr foregroundWindowBeforeTaskClick = IntPtr.Zero;
        private readonly WinApi.LowLevelKeyboardProc keyboardProc;
        private IntPtr keyboardHookHandle = IntPtr.Zero;
        private bool winKeyPressed;
        private bool nonWinKeyPressedWhileWinHeld;
        private readonly VolumePopupForm volumePopup;
        private readonly TimeDetailsForm timeDetailsPopup;
        private readonly PictureBox customVolumeIcon;
        private FileSystemWatcher? quickLaunchWatcher;
        private readonly System.Windows.Forms.Timer quickLaunchRefreshDebounceTimer;
        private string? quickLaunchContextShortcutPath;
        private int startMenuIconSize = DefaultStartMenuIconSize;
        private int taskIconSize = DefaultTaskIconSize;
        private string taskbarFontName = "MS Sans Serif";
        private float taskbarFontSize = 8.25f;
        private bool lazyLoadProgramsSubmenu = true;
        private bool playVolumeFeedbackSound = true;
        private int startMenuSubmenuOpenDelayMs = 200;
        private bool autoHideTaskbar;
        private TaskbarButtonStyle taskbarButtonStyle = TaskbarButtonStyle.Classic;
        private int taskbarBevelSize = 1;
        private Color taskbarBaseColor = Color.FromArgb(192, 192, 192);
        private Color taskbarLightColor = Color.FromArgb(255, 255, 255);
        private Color taskbarDarkColor = Color.FromArgb(128, 128, 128);
        private bool taskbarLocked;
        private int taskbarRows = 1;
        private bool isTaskbarResizeDragging;
        private readonly Dictionary<string, ProgramFolderCacheEntry> programFolderCache = new(StringComparer.OrdinalIgnoreCase);
        private sealed record ProgramFolderCacheEntry(List<string> Directories, List<string> Files);
        private readonly StringBuilder startMenuSearchQuery = new();
        private DateTime startMenuSearchLastInputUtc = DateTime.MinValue;
        private readonly List<ToolStripItem> startMenuSearchResultItems = new();
        private ToolStripTextBox? startMenuSearchTextBox;
        private readonly List<ProgramSearchEntry> programSearchEntriesCache = new();
        private readonly HashSet<ToolStripMenuItem> populatedSubmenusThisSession = new();
        private sealed record ProgramSearchEntry(string DisplayName, string LaunchPath, bool IsShellApp);
        private enum TaskbarButtonStyle
        {
            Classic,
            Win98
        }

        private sealed class Win9xColorTable : ProfessionalColorTable
        {
            public override Color ToolStripDropDownBackground => Color.FromArgb(192, 192, 192);
            public override Color ImageMarginGradientBegin => Color.FromArgb(192, 192, 192);
            public override Color ImageMarginGradientMiddle => Color.FromArgb(192, 192, 192);
            public override Color ImageMarginGradientEnd => Color.FromArgb(192, 192, 192);
            public override Color MenuItemSelected => Color.FromArgb(0, 0, 128);
            public override Color MenuItemSelectedGradientBegin => Color.FromArgb(0, 0, 128);
            public override Color MenuItemSelectedGradientEnd => Color.FromArgb(0, 0, 128);
            public override Color MenuItemBorder => Color.FromArgb(0, 0, 128);
            public override Color MenuBorder => Color.FromArgb(128, 128, 128);
            public override Color SeparatorDark => Color.FromArgb(128, 128, 128);
            public override Color SeparatorLight => Color.FromArgb(255, 255, 255);
        }

        private sealed class Win9xMenuRendererImpl : ToolStripProfessionalRenderer
        {
            private static readonly Color MenuBase = Color.FromArgb(192, 192, 192);
            private static readonly Color MenuBlue = Color.FromArgb(0, 0, 128);
            private static readonly Color MenuLight = Color.FromArgb(255, 255, 255);
            private static readonly Color MenuDark = Color.FromArgb(128, 128, 128);

            public Win9xMenuRendererImpl() : base(new Win9xColorTable())
            {
                RoundedEdges = false;
            }

            protected override void OnRenderToolStripBackground(ToolStripRenderEventArgs e)
            {
                using var brush = new SolidBrush(MenuBase);
                e.Graphics.FillRectangle(brush, e.AffectedBounds);
            }

            protected override void OnRenderToolStripBorder(ToolStripRenderEventArgs e)
            {
                var rect = new Rectangle(0, 0, e.ToolStrip.Width - 1, e.ToolStrip.Height - 1);
                if (rect.Width <= 1 || rect.Height <= 1)
                {
                    return;
                }

                using var darkPen = new Pen(MenuDark);
                using var lightPen = new Pen(MenuLight);

                // Raised 3D edge (top/left light, right/bottom dark)
                var left = rect.Left;
                var top = rect.Top;
                var right = rect.Right;
                var bottom = rect.Bottom;

                e.Graphics.DrawLine(lightPen, left, top, right, top);
                e.Graphics.DrawLine(lightPen, left, top, left, bottom);
                e.Graphics.DrawLine(darkPen, right, top, right, bottom);
                e.Graphics.DrawLine(darkPen, left, bottom, right, bottom);
            }

            protected override void OnRenderMenuItemBackground(ToolStripItemRenderEventArgs e)
            {
                var rect = new Rectangle(Point.Empty, e.Item.Size);

                using var backBrush = new SolidBrush(MenuBase);
                e.Graphics.FillRectangle(backBrush, rect);

                if (e.Item.Selected && e.Item.Enabled)
                {
                    var highlightRect = new Rectangle(1, 1, Math.Max(1, e.Item.Width - 2), Math.Max(1, e.Item.Height - 2));
                    using var highlightBrush = new SolidBrush(MenuBlue);
                    e.Graphics.FillRectangle(highlightBrush, highlightRect);
                }
            }

            protected override void OnRenderItemText(ToolStripItemTextRenderEventArgs e)
            {
                var text = e.Text ?? string.Empty;
                var font = e.TextFont ?? SystemFonts.MenuFont;
                var flags = TextFormatFlags.Left | TextFormatFlags.VerticalCenter;

                if (!e.Item.Enabled)
                {
                    var shadowRect = new Rectangle(e.TextRectangle.X + 1, e.TextRectangle.Y + 1, e.TextRectangle.Width, e.TextRectangle.Height);
                    TextRenderer.DrawText(e.Graphics, text, font, shadowRect, MenuLight, flags);
                    TextRenderer.DrawText(e.Graphics, text, font, e.TextRectangle, MenuDark, flags);
                    return;
                }

                var textColor = e.Item.Selected ? Color.White : Color.Black;
                TextRenderer.DrawText(e.Graphics, text, font, e.TextRectangle, textColor, flags);
            }

            protected override void OnRenderSeparator(ToolStripSeparatorRenderEventArgs e)
            {
                var y = e.Item.ContentRectangle.Top + (e.Item.ContentRectangle.Height / 2) - 1;
                var left = e.Item.ContentRectangle.Left + 2;
                var right = e.Item.ContentRectangle.Right - 2;

                using var darkPen = new Pen(MenuDark);
                using var lightPen = new Pen(MenuLight);
                e.Graphics.DrawLine(darkPen, left, y, right, y);
                e.Graphics.DrawLine(lightPen, left, y + 1, right, y + 1);
            }

            protected override void OnRenderArrow(ToolStripArrowRenderEventArgs e)
            {
                if (e.Item == null || e.ArrowRectangle.IsEmpty)
                {
                    return;
                }

                var selected = e.Item.Selected && e.Item.Enabled;
                var arrowColor = selected ? Color.White : Color.Black;

                using var brush = new SolidBrush(arrowColor);
                var rect = e.ArrowRectangle;
                var midX = rect.X + rect.Width / 2;
                var midY = rect.Y + rect.Height / 2;
                var arrowSize = Math.Min(rect.Width, rect.Height) / 2;

                // Draw a right-pointing triangle
                var points = new[]
                {
                    new Point(midX - arrowSize / 2, midY - arrowSize / 2),
                    new Point(midX - arrowSize / 2, midY + arrowSize / 2),
                    new Point(midX + arrowSize / 2, midY)
                };

                e.Graphics.FillPolygon(brush, points);
            }
        }

        private int TaskbarButtonHeight => TaskbarRowHeight - TaskbarButtonVerticalPadding - TaskbarBottomPadding - 1;

        private int QuickLaunchButtonSize => TaskbarButtonHeight;

        private int ExpandedTaskbarHeight => taskbarRows * TaskbarRowHeight;

        private int ResizeGripHeight => Math.Max(ResizeGripMinHeight, taskbarBevelSize);

        private int TotalTaskbarHeight => ExpandedTaskbarHeight + ResizeGripHeight;

        public RetroTaskbarForm()
        {
            instance = this;
            shellManager = new ShellManager();

            var settings = TaskbarSettingsStore.Load();
            startMenuIconSize = Math.Clamp(settings.StartMenuIconSize, 16, 32);
            taskIconSize = Math.Clamp(settings.TaskIconSize, 16, 32);
            lazyLoadProgramsSubmenu = settings.LazyLoadProgramsSubmenu;
            playVolumeFeedbackSound = settings.PlayVolumeFeedbackSound;
            startMenuSubmenuOpenDelayMs = Math.Clamp(settings.StartMenuSubmenuOpenDelayMs, 0, 1500);
            autoHideTaskbar = settings.AutoHideTaskbar;
            taskbarFontName = string.IsNullOrWhiteSpace(settings.TaskbarFontName) ? "MS Sans Serif" : settings.TaskbarFontName;
            taskbarFontSize = Math.Clamp(settings.TaskbarFontSize, 7f, 16f);
            taskbarLocked = settings.TaskbarLocked;
            taskbarRows = Math.Clamp(settings.TaskbarRows, 1, 3);
            taskbarBevelSize = Math.Clamp(settings.TaskbarBevelSize, 1, 4);
            taskbarBaseColor = ParseColorOrDefault(settings.TaskbarBaseColor, Color.FromArgb(192, 192, 192));
            taskbarLightColor = ParseColorOrDefault(settings.TaskbarLightColor, Color.FromArgb(255, 255, 255));
            taskbarDarkColor = ParseColorOrDefault(settings.TaskbarDarkColor, Color.FromArgb(128, 128, 128));
            if (Enum.TryParse<TaskbarButtonStyle>(settings.TaskbarButtonStyle, ignoreCase: true, out var parsedStyle))
            {
                taskbarButtonStyle = parsedStyle;
            }
            else if (string.Equals(settings.TaskbarButtonStyle, "Win98 Thick", StringComparison.OrdinalIgnoreCase))
            {
                taskbarButtonStyle = TaskbarButtonStyle.Win98;
                taskbarBevelSize = Math.Max(taskbarBevelSize, 2);
            }
            else if (string.Equals(settings.TaskbarButtonStyle, "Flat", StringComparison.OrdinalIgnoreCase) ||
                     string.Equals(settings.TaskbarButtonStyle, "Borderless", StringComparison.OrdinalIgnoreCase))
            {
                taskbarButtonStyle = TaskbarButtonStyle.Win98;
            }

            FormBorderStyle = FormBorderStyle.None;
            ShowIcon = false;
            ShowInTaskbar = false;
            TopMost = true;
            StartPosition = FormStartPosition.Manual;
            BackColor = taskbarBaseColor;
            Padding = Padding.Empty;
            normalFont = CreateTaskbarFont(FontStyle.Regular);
            activeFont = CreateTaskbarFont(FontStyle.Bold);
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

            resizeGripPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = ResizeGripHeight,
                BackColor = taskbarBaseColor,
                Cursor = Cursors.SizeNS,
                Visible = true
            };
            resizeGripPanel.Paint += ResizeGripPanel_Paint;

            contentPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = taskbarBaseColor,
                Padding = new Padding(0, TaskbarButtonVerticalPadding, 0, TaskbarBottomPadding)
            };

            startHostPanel = new Panel
            {
                Dock = DockStyle.Left,
                Width = 74,
                BackColor = taskbarBaseColor,
                Padding = new Padding(0, 0, TaskbarButtonGap, 0)
            };

            btnStart = new Button
            {
                Text = "Start",
                Width = 72,
                Height = TaskbarButtonHeight,
                Dock = DockStyle.Top,
                Font = activeFont,
                TabStop = false,
                Margin = Padding.Empty,
                TextAlign = ContentAlignment.MiddleLeft,
                TextImageRelation = TextImageRelation.ImageBeforeText,
                ImageAlign = ContentAlignment.MiddleLeft,
                Image = AddIconMargin(ResizeIconImage(startButtonMenuIcon, startMenuIconSize), 2, 1)
            };
            ApplyTaskbarButtonStyle(btnStart);
            AttachNoFocusCue(btnStart);
            btnStart.MouseDown += BtnStart_MouseDown;
            btnStart.Click += BtnStart_Click;

            lblClock = new Label
            {
                Dock = DockStyle.Right,
                Width = 72,
                TextAlign = ContentAlignment.MiddleCenter,
                BorderStyle = BorderStyle.None,
                BackColor = taskbarBaseColor,
                Margin = Padding.Empty,
                Font = normalFont
            };

            quickLaunchPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Left,
                Width = 0,
                WrapContents = false,
                AutoScroll = false,
                AllowDrop = true,
                Margin = Padding.Empty,
                Padding = new Padding(0, 0, TaskbarButtonGap, 0),
                BackColor = taskbarBaseColor,
                BorderStyle = BorderStyle.None
            };

            notificationAreaPanel = new Panel
            {
                Dock = DockStyle.Right,
                Width = 172,
                BackColor = taskbarBaseColor,
                BorderStyle = BorderStyle.None,
                Padding = new Padding(1, 1, 1, 1)
            };
            notificationAreaPanel.Paint += NotificationAreaPanel_Paint;

            trayIconsPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Left,
                Width = 96,
                WrapContents = false,
                AutoScroll = false,
                Margin = Padding.Empty,
                Padding = new Padding(2, 4, 2, 0),
                BackColor = taskbarBaseColor,
                BorderStyle = BorderStyle.None
            };

            customVolumeIcon = new PictureBox
            {
                Width = 18,
                Height = 18,
                Margin = new Padding(1, 0, 1, 0),
                SizeMode = PictureBoxSizeMode.CenterImage,
                Cursor = Cursors.Default,
                Image = ResizeIconImage(customVolumeIconImage, GetScaledSmallIconSize())
            };
            customVolumeIcon.MouseUp += CustomVolumeIcon_MouseUp;

            taskButtonsPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                WrapContents = false,
                AutoScroll = false,
                Margin = Padding.Empty,
                Padding = new Padding(0, 0, TaskbarButtonGap, 0),
                BackColor = taskbarBaseColor
            };

            startMenu = BuildStartMenu();
            taskWindowMenu = BuildTaskWindowMenu();
            taskbarContextMenu = BuildTaskbarContextMenu();
            clockContextMenu = BuildClockContextMenu();
            quickLaunchItemContextMenu = BuildQuickLaunchItemContextMenu();
            taskToolTip = new ToolTip();
            trayToolTip = new ToolTip
            {
                ShowAlways = true,
                AutomaticDelay = 300,
                ReshowDelay = 100,
                AutoPopDelay = 8000
            };
            quickLaunchToolTip = new ToolTip();
            keyboardProc = KeyboardHookCallback;
            volumePopup = new VolumePopupForm();
            volumePopup.PlayFeedbackSoundOnMouseUp = playVolumeFeedbackSound;
            timeDetailsPopup = new TimeDetailsForm();
            LoadQuickLaunchButtons();

            startHostPanel.Controls.Add(btnStart);
            notificationAreaPanel.Controls.Add(lblClock);
            notificationAreaPanel.Controls.Add(trayIconsPanel);
            contentPanel.Controls.Add(taskButtonsPanel);
            contentPanel.Controls.Add(notificationAreaPanel);
            contentPanel.Controls.Add(quickLaunchPanel);
            contentPanel.Controls.Add(startHostPanel);

            Controls.Add(contentPanel);
            Controls.Add(resizeGripPanel);
            trayIconsPanel.Controls.Add(customVolumeIcon);
            trayToolTip.SetToolTip(customVolumeIcon, "Volume");
            UpdateResizeUiState();

            MouseUp += TaskbarBackground_MouseUp;
            MouseDown += TaskbarResize_MouseDown;
            MouseMove += TaskbarResize_MouseMove;
            MouseUp += TaskbarResize_MouseUp;
            resizeGripPanel.MouseDown += TaskbarResize_MouseDown;
            resizeGripPanel.MouseMove += TaskbarResize_MouseMove;
            resizeGripPanel.MouseUp += TaskbarResize_MouseUp;
            taskButtonsPanel.MouseUp += TaskButtonsPanel_MouseUp;
            taskButtonsPanel.MouseDown += ChildTaskbarResize_MouseDown;
            taskButtonsPanel.MouseMove += ChildTaskbarResize_MouseMove;
            taskButtonsPanel.MouseUp += ChildTaskbarResize_MouseUp;
            lblClock.MouseUp += LblClock_MouseUp;
            lblClock.MouseClick += LblClock_MouseClick;
            lblClock.MouseDown += ChildTaskbarResize_MouseDown;
            lblClock.MouseMove += ChildTaskbarResize_MouseMove;
            lblClock.MouseUp += ChildTaskbarResize_MouseUp;
            quickLaunchPanel.MouseDown += ChildTaskbarResize_MouseDown;
            quickLaunchPanel.MouseMove += ChildTaskbarResize_MouseMove;
            quickLaunchPanel.MouseUp += ChildTaskbarResize_MouseUp;
            trayIconsPanel.MouseDown += ChildTaskbarResize_MouseDown;
            trayIconsPanel.MouseMove += ChildTaskbarResize_MouseMove;
            trayIconsPanel.MouseUp += ChildTaskbarResize_MouseUp;
            quickLaunchPanel.DragEnter += QuickLaunchPanel_DragEnter;
            quickLaunchPanel.DragDrop += QuickLaunchPanel_DragDrop;

            refreshTimer = new System.Windows.Forms.Timer { Interval = RefreshIntervalMs };
            refreshTimer.Tick += (_, _) => RefreshTaskbar();
            quickLaunchRefreshDebounceTimer = new System.Windows.Forms.Timer { Interval = 300 };
            quickLaunchRefreshDebounceTimer.Tick += (_, _) =>
            {
                quickLaunchRefreshDebounceTimer.Stop();
                LoadQuickLaunchButtons();
            };

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
            InitializeQuickLaunchWatcher();
            SetTaskbarBounds();
            RefreshTaskbar();
            refreshTimer.Start();
        }

        private void RetroTaskbarForm_FormClosed(object? sender, FormClosedEventArgs e)
        {
            instance = null;
            refreshTimer.Stop();
            refreshTimer.Dispose();
            quickLaunchRefreshDebounceTimer.Stop();
            quickLaunchRefreshDebounceTimer.Dispose();
            SaveTaskbarSettings();

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
                    quickLaunchButton.MouseUp -= QuickLaunchButton_MouseUp;
                    quickLaunchButton.Image?.Dispose();
                }

                control.Dispose();
            }

            quickLaunchPanel.DragEnter -= QuickLaunchPanel_DragEnter;
            quickLaunchPanel.DragDrop -= QuickLaunchPanel_DragDrop;

            btnStart.Image?.Dispose();
            startMenu.Dispose();
            taskWindowMenu.Dispose();
            taskbarContextMenu.Dispose();
            clockContextMenu.Dispose();
            quickLaunchItemContextMenu.Dispose();
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
            DisposeQuickLaunchWatcher();
            TeardownTrayService();

            TrySetExplorerTaskbarHidden(false);
            shellManager.AppBarManager.SignalGracefulShutdown();
            shellManager.Dispose();
        }

        private void RetroTaskbarForm_Paint(object? sender, PaintEventArgs e)
        {
            // Top bevel is rendered by resizeGripPanel.
        }

        private ContextMenuStrip BuildStartMenu()
        {
            var menu = new ContextMenuStrip
            {
                ShowImageMargin = true,
                Font = normalFont,
                ImageScalingSize = new Size(startMenuIconSize, startMenuIconSize)
            };
            ApplyWin9xContextMenuStyle(menu);

            startMenuSearchTextBox = new ToolStripTextBox
            {
                AutoSize = false,
                Width = 220,
                ToolTipText = "Type to search programs",
                ReadOnly = true
            };
            menu.Items.Add(startMenuSearchTextBox);
            menu.Items.Add(new ToolStripSeparator());

            menu.Items.Add(CreateMenuItem("File Explorer", explorerMenuIcon, (_, _) => BringExplorerToFront()));

            var programsItem = CreateMenuItem("Programs", programsMenuIcon, null);
            if (lazyLoadProgramsSubmenu)
            {
                var programsLoadingItem = CreateMenuItem("(Loading...)", null, null);
                programsLoadingItem.Enabled = false;
                programsItem.DropDownItems.Add(programsLoadingItem);
                programsItem.DropDownOpening += (_, _) => PopulateSubmenuWithOptionalDelay(programsItem, () => PopulateProgramsMenu(programsItem.DropDownItems));
            }
            else
            {
                PopulateProgramsMenu(programsItem.DropDownItems);
            }
            menu.Items.Add(programsItem);

            var windowsAppsItem = CreateMenuItem("Windows Apps", SystemIcons.Application.ToBitmap(), null);
            var windowsAppsLoadingItem = CreateMenuItem("(Loading...)", null, null);
            windowsAppsLoadingItem.Enabled = false;
            windowsAppsItem.DropDownItems.Add(windowsAppsLoadingItem);
            windowsAppsItem.DropDownOpening += (_, _) => PopulateSubmenuWithOptionalDelay(windowsAppsItem, () => PopulateWindowsAppsMenu(windowsAppsItem.DropDownItems));
            menu.Items.Add(windowsAppsItem);

            var documentsItem = CreateMenuItem("Documents", documentsMenuIcon, null);
            var documentsLoadingItem = CreateMenuItem("(Loading...)", null, null);
            documentsLoadingItem.Enabled = false;
            documentsItem.DropDownItems.Add(documentsLoadingItem);
            documentsItem.DropDownOpening += (_, _) => PopulateSubmenuWithOptionalDelay(documentsItem, () => PopulateDocumentsMenu(documentsItem.DropDownItems));
            menu.Items.Add(documentsItem);

            var settingsItem = CreateMenuItem("Settings", settingsMenuIcon, null);
            settingsItem.DropDownItems.Add(CreateMenuItem("Control Panel", SystemIcons.Shield.ToBitmap(), (_, _) => LaunchProcess("control.exe")));
            settingsItem.DropDownItems.Add(CreateMenuItem("Printers", LoadMenuAsset("Prints Folder.png", SystemIcons.Application.ToBitmap()), (_, _) => LaunchProcess("control.exe", "printers")));
            settingsItem.DropDownItems.Add(CreateMenuItem("Network Connections", LoadMenuAsset("Internet Network.png", SystemIcons.Application.ToBitmap()), (_, _) => OpenNetworkConnectionsSettings()));
            settingsItem.DropDownItems.Add(CreateMenuItem("Taskbar and Start Menu...", LoadMenuAsset("Taskbar and Start Menu.png", SystemIcons.Application.ToBitmap()), (_, _) => OpenTaskbarSettings()));
            settingsItem.DropDownItems.Add(CreateMenuItem("Folder Options...", LoadMenuAsset("Folder Open.png", SystemIcons.Application.ToBitmap()), (_, _) => OpenFolderOptionsSettings()));
            menu.Items.Add(settingsItem);

            menu.Items.Add(CreateMenuItem("Run...", runMenuIcon, (_, _) => LaunchProcess("rundll32.exe", "shell32.dll,#61")));
            menu.Items.Add(CreateMenuItem("Invalidate Start Menu Cache", null, (_, _) => InvalidateStartMenuCache()));
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add(CreateMenuItem("Shut Down...", SystemIcons.Error.ToBitmap(), (_, _) => ShowShutdownPrompt()));

            menu.Closed += (_, _) =>
            {
                lastStartMenuClosedAtUtc = DateTime.UtcNow;
                ResetStartMenuSearch();
                ResetStartMenuSubmenuSessionCache();
                btnStart.Invalidate();
            };

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

        private void AttachStartMenuShellContextMenu(ToolStripMenuItem item, string fileSystemPath)
        {
            item.MouseUp += (_, e) =>
            {
                if (e.Button != MouseButtons.Right)
                {
                    return;
                }

                ShowShellContextMenuForPath(fileSystemPath, Cursor.Position);
            };
        }

        private void ShowShellContextMenuForPath(string filePath, Point screenPoint)
        {
            IntPtr pidl = IntPtr.Zero;
            IntPtr parentFolder = IntPtr.Zero;
            IntPtr contextMenu = IntPtr.Zero;
            IntPtr menu = IntPtr.Zero;

            try
            {
                int result = WinApi.SHParseDisplayName(filePath, IntPtr.Zero, out pidl, 0, out _);
                if (result != 0 || pidl == IntPtr.Zero)
                {
                    return;
                }

                IntPtr pidlLast;
                result = WinApi.SHBindToParent(pidl, WinApi.IID_IShellFolder, out parentFolder, out pidlLast);
                if (result != 0 || parentFolder == IntPtr.Zero)
                {
                    return;
                }

                var shellFolder = Marshal.GetObjectForIUnknown(parentFolder) as WinApi.IShellFolder;
                if (shellFolder == null)
                {
                    return;
                }

                IntPtr[] pidlArray = { pidlLast };
                var contextMenuGuid = WinApi.IID_IContextMenu;
                result = shellFolder.GetUIObjectOf(Handle, 1, pidlArray, ref contextMenuGuid, IntPtr.Zero, out contextMenu);
                if (result != 0 || contextMenu == IntPtr.Zero)
                {
                    return;
                }

                var contextMenuInterface = Marshal.GetObjectForIUnknown(contextMenu) as WinApi.IContextMenu;
                if (contextMenuInterface == null)
                {
                    return;
                }

                menu = WinApi.CreatePopupMenu();
                if (menu == IntPtr.Zero)
                {
                    return;
                }

                result = contextMenuInterface.QueryContextMenu(menu, 0, 1, 0x7FFF, 0x10);
                if (result < 0)
                {
                    return;
                }

                uint command = (uint)WinApi.TrackPopupMenuEx(menu, WinApi.TPM_RETURNCMD | WinApi.TPM_LEFTBUTTON, screenPoint.X, screenPoint.Y, Handle, IntPtr.Zero);
                if (command == 0)
                {
                    return;
                }

                var commandInfo = new WinApi.CMINVOKECOMMANDINFO
                {
                    cbSize = Marshal.SizeOf(typeof(WinApi.CMINVOKECOMMANDINFO)),
                    fMask = 0,
                    hwnd = Handle,
                    lpVerb = new IntPtr(command - 1),
                    lpParameters = IntPtr.Zero,
                    lpDirectory = IntPtr.Zero,
                    nShow = 1,
                    dwHotKey = 0,
                    hIcon = IntPtr.Zero
                };

                IntPtr commandInfoPtr = Marshal.AllocHGlobal(Marshal.SizeOf(commandInfo));
                try
                {
                    Marshal.StructureToPtr(commandInfo, commandInfoPtr, false);
                    contextMenuInterface.InvokeCommand(commandInfoPtr);
                }
                finally
                {
                    Marshal.FreeHGlobal(commandInfoPtr);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to show shell context menu for '{filePath}': {ex.Message}");
            }
            finally
            {
                if (menu != IntPtr.Zero)
                {
                    WinApi.DestroyMenu(menu);
                }

                if (contextMenu != IntPtr.Zero)
                {
                    Marshal.Release(contextMenu);
                }

                if (parentFolder != IntPtr.Zero)
                {
                    Marshal.Release(parentFolder);
                }

                if (pidl != IntPtr.Zero)
                {
                    WinApi.CoTaskMemFree(pidl);
                }
            }
        }

        private void PopulateProgramsMenu(ToolStripItemCollection items)
        {
            EnsureProgramMenuCacheFresh();
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

            ApplyTwoColumnLayoutIfNeeded(items);
        }

        private void AddProgramEntries(ToolStripItemCollection items, string folderPath, int depth)
        {
            if (depth > ProgramMenuMaxDepth)
            {
                return;
            }

            var entries = GetProgramFolderEntries(folderPath);
            if (entries == null)
            {
                return;
            }

            foreach (var directory in entries.Directories)
            {
                var subMenu = CreateProgramFolderMenuItem(Path.GetFileName(directory), directory, depth + 1, lazyLoadProgramsSubmenu);
                if (subMenu != null)
                {
                    AttachStartMenuShellContextMenu(subMenu, directory);
                    items.Add(subMenu);
                }
            }

            foreach (var file in entries.Files)
            {
                var name = Path.GetFileNameWithoutExtension(file);
                var item = CreateMenuItem(name, GetSmallShellIcon(file, isDirectory: false), (_, _) => LaunchProcess(file));
                AttachStartMenuShellContextMenu(item, file);
                items.Add(item);
            }
        }

        private void EnsureProgramMenuCacheFresh()
        {
            // cache is kept until manually invalidated
        }

        private void InvalidateProgramMenuCache()
        {
            programFolderCache.Clear();
            programSearchEntriesCache.Clear();
        }

        private void InvalidateStartMenuCache()
        {
            InvalidateProgramMenuCache();
            ResetStartMenuSubmenuSessionCache();
            UpdateStartMenuSearchResultsFromCurrentQuery();
        }

        private ProgramFolderCacheEntry? GetProgramFolderEntries(string folderPath)
        {
            EnsureProgramMenuCacheFresh();

            if (programFolderCache.TryGetValue(folderPath, out var cached))
            {
                return cached;
            }

            try
            {
                var directories = Directory.GetDirectories(folderPath).OrderBy(Path.GetFileName).ToList();
                var files = ProgramLaunchPatterns
                    .SelectMany(pattern => Directory.GetFiles(folderPath, pattern))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(Path.GetFileName)
                    .ToList();

                var entry = new ProgramFolderCacheEntry(directories, files);
                programFolderCache[folderPath] = entry;
                return entry;
            }
            catch
            {
                return null;
            }
        }

        private ToolStripMenuItem? CreateProgramFolderMenuItem(string displayName, string folderPath, int depth, bool lazyLoadChildren)
        {
            if (depth > ProgramMenuMaxDepth || !Directory.Exists(folderPath))
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

            PopulateSubmenuWithOptionalDelay(item, () =>
            {
                item.DropDownItems.Clear();
                AddProgramEntries(item.DropDownItems, node.FolderPath, node.Depth);

                if (item.DropDownItems.Count == 0)
                {
                    var empty = CreateMenuItem("(Empty)", null, null);
                    empty.Enabled = false;
                    item.DropDownItems.Add(empty);
                }

                ApplyTwoColumnLayoutIfNeeded(item.DropDownItems);
            });
        }

        private void PopulateSubmenuWithOptionalDelay(ToolStripMenuItem item, Action populateAction)
        {
            if (populatedSubmenusThisSession.Contains(item))
            {
                return;
            }

            if (startMenuSubmenuOpenDelayMs <= 0)
            {
                populateAction();
                populatedSubmenusThisSession.Add(item);
                return;
            }

            item.DropDownItems.Clear();
            var loadingItem = CreateMenuItem("(Loading...)", null, null);
            loadingItem.Enabled = false;
            item.DropDownItems.Add(loadingItem);

            var timer = new System.Windows.Forms.Timer { Interval = startMenuSubmenuOpenDelayMs };
            timer.Tick += (_, _) =>
            {
                timer.Stop();
                timer.Dispose();

                if (!item.Selected || !startMenu.Visible)
                {
                    return;
                }

                populateAction();
                populatedSubmenusThisSession.Add(item);
                if (!item.DropDown.Visible)
                {
                    item.ShowDropDown();
                }
            };
            timer.Start();
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

        private void PopulateWindowsAppsMenu(ToolStripItemCollection items)
        {
            items.Clear();

            var apps = GetProgramSearchEntries()
                .Where(entry => entry.IsShellApp)
                .OrderBy(entry => entry.DisplayName, StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            if (apps.Count == 0)
            {
                var emptyItem = CreateMenuItem("(No Windows Apps found)", null, null);
                emptyItem.Enabled = false;
                items.Add(emptyItem);
                return;
            }

            foreach (var app in apps)
            {
                items.Add(CreateMenuItem(app.DisplayName, GetProgramSearchEntryIcon(app), (_, _) => LaunchProgramSearchEntry(app)));
            }

            ApplyTwoColumnLayoutIfNeeded(items);
        }

        private void ApplyTwoColumnLayoutIfNeeded(ToolStripItemCollection items)
        {
            if (items.Count == 0)
            {
                return;
            }

            var owner = items.Cast<ToolStripItem>().FirstOrDefault()?.Owner;
            if (owner is not ToolStripDropDown dropDown)
            {
                return;
            }

            var visibleItemCount = items.Cast<ToolStripItem>().Count(item => item.Available && item is not ToolStripSeparator);
            if (visibleItemCount < MultiColumnThreshold)
            {
                dropDown.LayoutStyle = ToolStripLayoutStyle.StackWithOverflow;
                dropDown.AutoSize = true;
                dropDown.MaximumSize = Size.Empty;
                foreach (ToolStripItem item in items)
                {
                    item.AutoSize = true;
                }

                return;
            }

            var itemHeight = 24;
            var maxHeight = Math.Max(240, (Screen.PrimaryScreen?.WorkingArea.Height ?? 768) - 120);
            var rowsPerColumn = Math.Max(1, (maxHeight - 8) / itemHeight);
            var columns = Math.Clamp((int)Math.Ceiling(visibleItemCount / (double)rowsPerColumn), 2, MaxSubmenuColumns);
            var rows = (int)Math.Ceiling(visibleItemCount / (double)columns);
            var targetHeight = Math.Min((rows * itemHeight) + 8, maxHeight);

            var maxMeasuredWidth = items.Cast<ToolStripItem>()
                .Where(item => item.Available && item is not ToolStripSeparator)
                .Select(item => TextRenderer.MeasureText(item.Text ?? string.Empty, item.Font ?? normalFont).Width)
                .DefaultIfEmpty(140)
                .Max();
            var imageSpace = startMenuIconSize + 26;
            var itemWidth = Math.Clamp(maxMeasuredWidth + imageSpace, 170, 360);
            var targetWidth = (itemWidth * columns) + 12;

            dropDown.LayoutStyle = ToolStripLayoutStyle.Flow;
            dropDown.AutoSize = false;
            dropDown.MaximumSize = new Size(targetWidth, targetHeight);
            dropDown.Size = new Size(targetWidth, targetHeight);

            if (dropDown.LayoutSettings is FlowLayoutSettings flowSettings)
            {
                flowSettings.FlowDirection = FlowDirection.TopDown;
                flowSettings.WrapContents = true;
            }

            foreach (ToolStripItem item in items)
            {
                if (item is ToolStripSeparator)
                {
                    item.AutoSize = true;
                    continue;
                }

                item.AutoSize = false;
                item.Width = itemWidth;
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

        private void ResetStartMenuSearch()
        {
            startMenuSearchQuery.Clear();
            startMenuSearchLastInputUtc = DateTime.MinValue;

            if (startMenuSearchTextBox != null)
            {
                startMenuSearchTextBox.Text = string.Empty;
            }

            ClearStartMenuSearchResultsItems();
            RepositionVisibleStartMenu();
        }

        private void ResetStartMenuSubmenuSessionCache()
        {
            populatedSubmenusThisSession.Clear();
        }

        private bool HandleStartMenuSearchKey(int vkCode)
        {
            if (!startMenu.Visible)
            {
                return false;
            }

            if (startMenuSearchTextBox == null)
            {
                return false;
            }

            var now = DateTime.UtcNow;

            if (vkCode == (int)Keys.Escape)
            {
                if (startMenuSearchQuery.Length > 0)
                {
                    ResetStartMenuSearch();
                    return true;
                }

                startMenu.Close(ToolStripDropDownCloseReason.Keyboard);
                return true;
            }

            if (vkCode == (int)Keys.Back)
            {
                if (startMenuSearchQuery.Length > 0)
                {
                    startMenuSearchQuery.Length--;
                    ApplySearchQueryToTextBox();
                    UpdateStartMenuSearchResultsFromCurrentQuery();
                }

                return true;
            }

            if (vkCode == (int)Keys.Return)
            {
                if (startMenuSearchQuery.Length == 0)
                {
                    return HandleStartMenuNavigationKey(vkCode);
                }

                var selectedResult = startMenuSearchResultItems
                    .OfType<ToolStripMenuItem>()
                    .FirstOrDefault(item => item.Selected && item.Tag is ProgramSearchEntry);

                if (selectedResult?.Tag is ProgramSearchEntry selectedEntry)
                {
                    LaunchProgramSearchEntry(selectedEntry);
                    startMenu.Close(ToolStripDropDownCloseReason.ItemClicked);
                    return true;
                }

                var topMatch = GetProgramSearchMatches(startMenuSearchQuery.ToString(), 1).FirstOrDefault();

                if (topMatch != null)
                {
                    LaunchProgramSearchEntry(topMatch);
                    startMenu.Close(ToolStripDropDownCloseReason.ItemClicked);
                }

                return true;
            }

            if (vkCode == (int)Keys.Down || vkCode == (int)Keys.Up)
            {
                if (startMenuSearchQuery.Length == 0)
                {
                    return HandleStartMenuNavigationKey(vkCode);
                }

                var selectableItems = startMenuSearchResultItems
                    .OfType<ToolStripMenuItem>()
                    .Where(item => item.Enabled && item.Tag is ProgramSearchEntry)
                    .ToList();

                if (selectableItems.Count == 0)
                {
                    return true;
                }

                var currentIndex = selectableItems.FindIndex(item => item.Selected);
                var nextIndex = vkCode == (int)Keys.Down
                    ? (currentIndex + 1 + selectableItems.Count) % selectableItems.Count
                    : (currentIndex - 1 + selectableItems.Count) % selectableItems.Count;

                selectableItems[nextIndex].Select();
                return true;
            }

            if (vkCode == (int)Keys.Right || vkCode == (int)Keys.Left || vkCode == (int)Keys.Tab)
            {
                return HandleStartMenuNavigationKey(vkCode);
            }

            var searchChar = TryConvertVirtualKeyToSearchChar(vkCode);
            if (searchChar == null)
            {
                return false;
            }

            if ((now - startMenuSearchLastInputUtc).TotalSeconds > StartMenuSearchResetSeconds)
            {
                startMenuSearchQuery.Clear();
            }

            startMenuSearchLastInputUtc = now;

            startMenuSearchQuery.Append(searchChar.Value);
            ApplySearchQueryToTextBox();
            UpdateStartMenuSearchResultsFromCurrentQuery();
            return true;
        }

        private bool HandleStartMenuNavigationKey(int vkCode)
        {
            if (!startMenu.Visible)
            {
                return false;
            }

            var openTopLevelDropDownHost = startMenu.Items
                .OfType<ToolStripMenuItem>()
                .FirstOrDefault(item => item.DropDown.Visible);
            if (openTopLevelDropDownHost != null)
            {
                return HandleOpenDropdownNavigation(openTopLevelDropDownHost, vkCode);
            }

            var navigableItems = startMenu.Items
                .OfType<ToolStripItem>()
                .Where(IsNavigableTopLevelMenuItem)
                .ToList();

            if (navigableItems.Count == 0)
            {
                return true;
            }

            var selectedIndex = navigableItems.FindIndex(item => item.Selected);

            if (vkCode == (int)Keys.Down || vkCode == (int)Keys.Tab)
            {
                var nextIndex = (selectedIndex + 1 + navigableItems.Count) % navigableItems.Count;
                navigableItems[nextIndex].Select();
                return true;
            }

            if (vkCode == (int)Keys.Up)
            {
                var nextIndex = (selectedIndex - 1 + navigableItems.Count) % navigableItems.Count;
                navigableItems[nextIndex].Select();
                return true;
            }

            if (vkCode == (int)Keys.Right || vkCode == (int)Keys.Return)
            {
                if (navigableItems.ElementAtOrDefault(Math.Max(selectedIndex, 0)) is not ToolStripMenuItem selectedMenuItem || !selectedMenuItem.Enabled)
                {
                    return true;
                }

                if (selectedMenuItem.HasDropDownItems)
                {
                    selectedMenuItem.ShowDropDown();
                    SelectFirstDropdownItem(selectedMenuItem);
                    return true;
                }

                if (vkCode == (int)Keys.Return)
                {
                    selectedMenuItem.PerformClick();
                }

                return true;
            }

            if (vkCode == (int)Keys.Left)
            {
                if (navigableItems.ElementAtOrDefault(Math.Max(selectedIndex, 0)) is ToolStripMenuItem selectedMenuItem && selectedMenuItem.DropDown.Visible)
                {
                    selectedMenuItem.HideDropDown();
                }

                return true;
            }

            return false;
        }

        private static bool HandleOpenDropdownNavigation(ToolStripMenuItem host, int vkCode)
        {
            var dropdownItems = host.DropDownItems
                .OfType<ToolStripItem>()
                .Where(item => item is not ToolStripSeparator && item.Available && item.Enabled)
                .ToList();

            if (dropdownItems.Count == 0)
            {
                return true;
            }

            var selectedIndex = dropdownItems.FindIndex(item => item.Selected);
            if (selectedIndex < 0)
            {
                dropdownItems[0].Select();
                selectedIndex = 0;
            }

            if (vkCode == (int)Keys.Down || vkCode == (int)Keys.Tab)
            {
                var nextIndex = (selectedIndex + 1 + dropdownItems.Count) % dropdownItems.Count;
                dropdownItems[nextIndex].Select();
                return true;
            }

            if (vkCode == (int)Keys.Up)
            {
                var nextIndex = (selectedIndex - 1 + dropdownItems.Count) % dropdownItems.Count;
                dropdownItems[nextIndex].Select();
                return true;
            }

            var selectedItem = dropdownItems[Math.Clamp(selectedIndex, 0, dropdownItems.Count - 1)];

            if (vkCode == (int)Keys.Right || vkCode == (int)Keys.Return)
            {
                if (selectedItem is ToolStripMenuItem selectedMenuItem)
                {
                    if (selectedMenuItem.HasDropDownItems)
                    {
                        selectedMenuItem.ShowDropDown();
                        SelectFirstDropdownItem(selectedMenuItem);
                    }
                    else if (vkCode == (int)Keys.Return)
                    {
                        selectedMenuItem.PerformClick();
                    }
                }

                return true;
            }

            if (vkCode == (int)Keys.Left || vkCode == (int)Keys.Escape)
            {
                host.HideDropDown();
                host.Select();
                return true;
            }

            return false;
        }

        private static bool IsNavigableTopLevelMenuItem(ToolStripItem item)
        {
            return item is not ToolStripSeparator and not ToolStripTextBox && item.Available && item.Enabled;
        }

        private static void SelectFirstDropdownItem(ToolStripMenuItem parent)
        {
            var first = parent.DropDownItems
                .OfType<ToolStripItem>()
                .FirstOrDefault(item => item is not ToolStripSeparator && item.Available && item.Enabled);
            first?.Select();
        }

        private void ApplySearchQueryToTextBox()
        {
            if (startMenuSearchTextBox == null)
            {
                return;
            }

            startMenuSearchTextBox.Text = startMenuSearchQuery.ToString();
            startMenuSearchTextBox.SelectionStart = startMenuSearchTextBox.TextLength;
        }

        private void UpdateStartMenuSearchResultsFromCurrentQuery()
        {
            UpdateStartMenuSearchResults(startMenuSearchQuery.ToString());
        }

        private void UpdateStartMenuSearchResults(string query)
        {
            ClearStartMenuSearchResultsItems();

            if (string.IsNullOrWhiteSpace(query))
            {
                RepositionVisibleStartMenu();
                return;
            }

            var insertIndex = Math.Min(2, startMenu.Items.Count);

            var matches = GetProgramSearchMatches(query, 12);
            if (matches.Count == 0)
            {
                var noMatchItem = CreateMenuItem($"No matches for \"{query}\"", null, null);
                noMatchItem.Enabled = false;
                startMenu.Items.Insert(insertIndex, noMatchItem);
                startMenuSearchResultItems.Add(noMatchItem);
            }
            else
            {
                for (var i = 0; i < matches.Count; i++)
                {
                    var match = matches[i];
                    var icon = GetProgramSearchEntryIcon(match);
                    var matchItem = CreateMenuItem(match.DisplayName, icon, (_, _) => LaunchProgramSearchEntry(match));
                    matchItem.Tag = match;
                    startMenu.Items.Insert(insertIndex + i, matchItem);
                    startMenuSearchResultItems.Add(matchItem);
                }
            }

            RepositionVisibleStartMenu();
        }

        private List<ProgramSearchEntry> GetProgramSearchMatches(string query, int maxCount)
        {
            var uniqueByName = new HashSet<string>(StringComparer.CurrentCultureIgnoreCase);

            return GetProgramSearchEntries()
                .Where(entry => entry.DisplayName.Contains(query, StringComparison.OrdinalIgnoreCase))
                .OrderBy(entry => !entry.DisplayName.StartsWith(query, StringComparison.OrdinalIgnoreCase))
                .ThenBy(entry => entry.DisplayName, StringComparer.CurrentCultureIgnoreCase)
                .Where(entry => uniqueByName.Add(entry.DisplayName.Trim()))
                .Take(maxCount)
                .ToList();
        }

        private void ClearStartMenuSearchResultsItems()
        {
            foreach (var item in startMenuSearchResultItems)
            {
                startMenu.Items.Remove(item);
                item.Dispose();
            }

            startMenuSearchResultItems.Clear();
        }

        private List<ProgramSearchEntry> GetProgramSearchEntries()
        {
            EnsureProgramMenuCacheFresh();

            if (programSearchEntriesCache.Count > 0)
            {
                return programSearchEntriesCache;
            }

            programSearchEntriesCache.Clear();
            var dedupKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var roots = new[]
            {
                Environment.GetFolderPath(Environment.SpecialFolder.Programs),
                Environment.GetFolderPath(Environment.SpecialFolder.CommonPrograms)
            };

            foreach (var root in roots.Distinct(StringComparer.OrdinalIgnoreCase))
            {
                if (!Directory.Exists(root))
                {
                    continue;
                }

                CollectProgramSearchEntries(root, depth: 0, dedupKeys);
            }

            CollectStartAppsSearchEntries(dedupKeys);

            return programSearchEntriesCache;
        }

        private void CollectProgramSearchEntries(string folderPath, int depth, HashSet<string> dedupKeys)
        {
            if (depth > ProgramMenuMaxDepth)
            {
                return;
            }

            var entries = GetProgramFolderEntries(folderPath);
            if (entries == null)
            {
                return;
            }

            foreach (var file in entries.Files)
            {
                var name = Path.GetFileNameWithoutExtension(file);
                if (string.IsNullOrWhiteSpace(name))
                {
                    continue;
                }

                var key = $"{name.Trim()}|{file}";
                if (!dedupKeys.Add(key))
                {
                    continue;
                }

                programSearchEntriesCache.Add(new ProgramSearchEntry(name, file, IsShellApp: false));
            }

            foreach (var directory in entries.Directories)
            {
                CollectProgramSearchEntries(directory, depth + 1, dedupKeys);
            }
        }

        private void CollectStartAppsSearchEntries(HashSet<string> dedupKeys)
        {
            try
            {
                using var process = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = "powershell.exe",
                        Arguments = "-NoProfile -ExecutionPolicy Bypass -Command \"Get-StartApps | ForEach-Object { '{0}|{1}' -f $_.Name, $_.AppID }\"",
                        UseShellExecute = false,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        CreateNoWindow = true
                    }
                };

                process.Start();
                var output = process.StandardOutput.ReadToEnd();
                process.WaitForExit(2000);

                if (string.IsNullOrWhiteSpace(output))
                {
                    return;
                }

                var lines = output.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var line in lines)
                {
                    var sepIndex = line.LastIndexOf('|');
                    if (sepIndex <= 0 || sepIndex >= line.Length - 1)
                    {
                        continue;
                    }

                    var name = line[..sepIndex].Trim();
                    var appId = line[(sepIndex + 1)..].Trim();
                    if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(appId))
                    {
                        continue;
                    }

                    var key = name.Trim() + "|" + appId;
                    if (!dedupKeys.Add(key))
                    {
                        continue;
                    }

                    programSearchEntriesCache.Add(new ProgramSearchEntry(name, appId, IsShellApp: true));
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to collect Start Apps: {ex.Message}");
            }
        }

        private static void LaunchProgramSearchEntry(ProgramSearchEntry entry)
        {
            if (entry.IsShellApp)
            {
                LaunchProcess("explorer.exe", $"shell:AppsFolder\\{entry.LaunchPath}");
                return;
            }

            LaunchProcess(entry.LaunchPath);
        }

        private Image? GetProgramSearchEntryIcon(ProgramSearchEntry entry)
        {
            if (!entry.IsShellApp)
            {
                return GetSmallShellIcon(entry.LaunchPath, isDirectory: false);
            }

            var shellPath = $"shell:AppsFolder\\{entry.LaunchPath}";
            return GetSmallShellIconFromShellPath(shellPath) ?? SystemIcons.Application.ToBitmap();
        }

        private static char? TryConvertVirtualKeyToSearchChar(int vkCode)
        {
            if (vkCode >= (int)Keys.A && vkCode <= (int)Keys.Z)
            {
                return (char)('a' + (vkCode - (int)Keys.A));
            }

            if (vkCode >= (int)Keys.D0 && vkCode <= (int)Keys.D9)
            {
                return (char)('0' + (vkCode - (int)Keys.D0));
            }

            if (vkCode >= (int)Keys.NumPad0 && vkCode <= (int)Keys.NumPad9)
            {
                return (char)('0' + (vkCode - (int)Keys.NumPad0));
            }

            return vkCode switch
            {
                (int)Keys.Space => ' ',
                (int)Keys.OemMinus => '-',
                (int)Keys.OemPeriod => '.',
                _ => null
            };
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

        private static Image? GetSmallShellIconFromShellPath(string shellPath)
        {
            IntPtr pidl = IntPtr.Zero;

            try
            {
                var parseResult = WinApi.SHParseDisplayName(shellPath, IntPtr.Zero, out pidl, 0, out _);
                if (parseResult != 0 || pidl == IntPtr.Zero)
                {
                    return null;
                }

                var shfi = new WinApi.SHFILEINFO();
                var flags = WinApi.SHGFI_PIDL | WinApi.SHGFI_ICON | WinApi.SHGFI_SMALLICON;
                WinApi.SHGetFileInfo(pidl, 0, ref shfi, (uint)Marshal.SizeOf(shfi), flags);

                if (shfi.hIcon == IntPtr.Zero)
                {
                    return null;
                }

                using var icon = Icon.FromHandle(shfi.hIcon);
                var bitmap = icon.ToBitmap();
                WinApi.DestroyIcon(shfi.hIcon);
                return new Bitmap(bitmap, new Size(16, 16));
            }
            finally
            {
                if (pidl != IntPtr.Zero)
                {
                    WinApi.CoTaskMemFree(pidl);
                }
            }
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
            ApplyWin9xContextMenuStyle(menu);

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
            ApplyWin9xContextMenuStyle(menu);

            menu.Items.Add("Refresh Quick Launch", null, (_, _) => LoadQuickLaunchButtons());
            menu.Items.Add(new ToolStripSeparator());
            var lockItem = new ToolStripMenuItem("Lock Taskbar", null, (_, _) => SetTaskbarLocked(true));
            var unlockItem = new ToolStripMenuItem("Unlock Taskbar", null, (_, _) => SetTaskbarLocked(false));
            menu.Items.Add(lockItem);
            menu.Items.Add(unlockItem);
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("Options...", null, (_, _) => ShowOptionsDialog());

            menu.Opening += (_, _) =>
            {
                lockItem.Enabled = !taskbarLocked;
                unlockItem.Enabled = taskbarLocked;
            };
            return menu;
        }

        private ContextMenuStrip BuildClockContextMenu()
        {
            var menu = new ContextMenuStrip
            {
                ShowImageMargin = false,
                Font = normalFont
            };
            ApplyWin9xContextMenuStyle(menu);

            menu.Items.Add("Open Clock...", null, (_, _) => ShowTimeDetailsModal());
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("Adjust Date/Time...", null, (_, _) => OpenDateTimeSettings());
            return menu;
        }

        private ContextMenuStrip BuildQuickLaunchItemContextMenu()
        {
            var menu = new ContextMenuStrip
            {
                ShowImageMargin = false,
                Font = normalFont
            };
            ApplyWin9xContextMenuStyle(menu);

            menu.Items.Add("Open", null, (_, _) => OpenQuickLaunchContextItem());
            menu.Items.Add("Open Quick Launch Folder", null, (_, _) => OpenQuickLaunchFolderInExplorerForm());
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("Remove from Quick Launch", null, (_, _) => RemoveQuickLaunchContextItem());
            menu.Items.Add("Properties", null, (_, _) => ShowQuickLaunchContextItemProperties());

            menu.Opening += (_, _) =>
            {
                var hasFile = !string.IsNullOrWhiteSpace(quickLaunchContextShortcutPath) && File.Exists(quickLaunchContextShortcutPath);
                menu.Items[0].Enabled = hasFile;
                menu.Items[3].Enabled = hasFile;
                menu.Items[4].Enabled = hasFile;
            };

            return menu;
        }

        private void InitializeQuickLaunchWatcher()
        {
            var quickLaunchFolder = GetQuickLaunchFolderPath();
            Directory.CreateDirectory(quickLaunchFolder);

            quickLaunchWatcher = new FileSystemWatcher(quickLaunchFolder)
            {
                IncludeSubdirectories = false,
                NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite | NotifyFilters.CreationTime,
                EnableRaisingEvents = true
            };

            quickLaunchWatcher.Changed += QuickLaunchWatcher_Changed;
            quickLaunchWatcher.Created += QuickLaunchWatcher_Changed;
            quickLaunchWatcher.Deleted += QuickLaunchWatcher_Changed;
            quickLaunchWatcher.Renamed += QuickLaunchWatcher_Renamed;
        }

        private static void ApplyWin9xContextMenuStyle(ContextMenuStrip menu)
        {
            menu.RenderMode = ToolStripRenderMode.Professional;
            menu.Renderer = Win9xMenuRenderer;
            menu.ShowImageMargin = menu.ShowImageMargin;
            menu.DropShadowEnabled = false;
        }

        private void DisposeQuickLaunchWatcher()
        {
            if (quickLaunchWatcher == null)
            {
                return;
            }

            quickLaunchWatcher.Changed -= QuickLaunchWatcher_Changed;
            quickLaunchWatcher.Created -= QuickLaunchWatcher_Changed;
            quickLaunchWatcher.Deleted -= QuickLaunchWatcher_Changed;
            quickLaunchWatcher.Renamed -= QuickLaunchWatcher_Renamed;
            quickLaunchWatcher.Dispose();
            quickLaunchWatcher = null;
        }

        private void QuickLaunchWatcher_Changed(object sender, FileSystemEventArgs e)
        {
            TriggerQuickLaunchRefresh();
        }

        private void QuickLaunchWatcher_Renamed(object sender, RenamedEventArgs e)
        {
            TriggerQuickLaunchRefresh();
        }

        private void TriggerQuickLaunchRefresh()
        {
            if (IsDisposed)
            {
                return;
            }

            if (InvokeRequired)
            {
                BeginInvoke(new Action(TriggerQuickLaunchRefresh));
                return;
            }

            quickLaunchRefreshDebounceTimer.Stop();
            quickLaunchRefreshDebounceTimer.Start();
        }

        private void SaveTaskbarSettings()
        {
            TaskbarSettingsStore.Save(new TaskbarSettings
            {
                StartMenuIconSize = startMenuIconSize,
                TaskIconSize = taskIconSize,
                LazyLoadProgramsSubmenu = lazyLoadProgramsSubmenu,
                PlayVolumeFeedbackSound = playVolumeFeedbackSound,
                StartMenuSubmenuOpenDelayMs = startMenuSubmenuOpenDelayMs,
                AutoHideTaskbar = autoHideTaskbar,
                TaskbarButtonStyle = taskbarButtonStyle.ToString(),
                TaskbarFontName = taskbarFontName,
                TaskbarFontSize = taskbarFontSize,
                TaskbarLocked = taskbarLocked,
                TaskbarRows = taskbarRows,
                TaskbarBaseColor = ColorTranslator.ToHtml(taskbarBaseColor),
                TaskbarLightColor = ColorTranslator.ToHtml(taskbarLightColor),
                TaskbarDarkColor = ColorTranslator.ToHtml(taskbarDarkColor),
                TaskbarBevelSize = taskbarBevelSize
            });
        }

        private Font CreateTaskbarFont(FontStyle style)
        {
            try
            {
                return new Font(taskbarFontName, taskbarFontSize, style, GraphicsUnit.Point);
            }
            catch
            {
                taskbarFontName = "MS Sans Serif";
                taskbarFontSize = 8.25f;
                return new Font(taskbarFontName, taskbarFontSize, style, GraphicsUnit.Point);
            }
        }

        private static Color ParseColorOrDefault(string? value, Color fallback)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return fallback;
            }

            try
            {
                return ColorTranslator.FromHtml(value);
            }
            catch
            {
                return fallback;
            }
        }

        private void RefreshTaskbarFonts()
        {
            normalFont.Dispose();
            activeFont.Dispose();
            normalFont = CreateTaskbarFont(FontStyle.Regular);
            activeFont = CreateTaskbarFont(FontStyle.Bold);

            Font = normalFont;
            lblClock.Font = normalFont;
            btnStart.Font = activeFont;
            btnStart.Height = TaskbarButtonHeight;

            taskWindowMenu.Font = normalFont;
            taskbarContextMenu.Font = normalFont;
            clockContextMenu.Font = normalFont;
            quickLaunchItemContextMenu.Font = normalFont;
            startMenu.Font = normalFont;
        }

        private void RefreshTaskbarColors()
        {
            BackColor = taskbarBaseColor;
            resizeGripPanel.BackColor = taskbarBaseColor;
            contentPanel.BackColor = taskbarBaseColor;
            startHostPanel.BackColor = taskbarBaseColor;
            taskButtonsPanel.BackColor = taskbarBaseColor;
            quickLaunchPanel.BackColor = taskbarBaseColor;
            trayIconsPanel.BackColor = taskbarBaseColor;
            notificationAreaPanel.BackColor = taskbarBaseColor;
            lblClock.BackColor = taskbarBaseColor;
            customVolumeIcon.BackColor = taskbarBaseColor;

            UpdateResizeUiState();
            notificationAreaPanel.Invalidate();
            resizeGripPanel.Invalidate();
        }

        private void ApplyTaskbarButtonStyle(Button button)
        {
            button.Paint -= TaskbarButtonWin98_Paint;

            switch (taskbarButtonStyle)
            {
                case TaskbarButtonStyle.Win98:
                    button.FlatStyle = FlatStyle.Flat;
                    button.FlatAppearance.BorderSize = 0;
                    button.FlatAppearance.BorderColor = taskbarBaseColor;
                    button.FlatAppearance.MouseOverBackColor = taskbarBaseColor;
                    button.FlatAppearance.MouseDownBackColor = taskbarBaseColor;
                    button.Paint += TaskbarButtonWin98_Paint;
                    button.UseVisualStyleBackColor = false;
                    button.BackColor = taskbarBaseColor;
                    break;
                default:
                    button.FlatStyle = FlatStyle.Standard;
                    button.UseVisualStyleBackColor = true;
                    break;
            }

            button.Invalidate();
        }

        private void TaskbarButtonWin98_Paint(object? sender, PaintEventArgs e)
        {
            if (sender is not Button button)
            {
                return;
            }

            var borderWidth = taskbarBevelSize;
            var rect = new Rectangle(0, 0, button.Width - 1, button.Height - 1);
            var isMousePressed = button.Capture && button.ClientRectangle.Contains(button.PointToClient(Cursor.Position));
            var isSelectedPressed = IsButtonSelectedPressed(button);
            var isPressed = isMousePressed || isSelectedPressed;

            if (isPressed)
            {
                ControlPaint.DrawBorder(e.Graphics, rect,
                    taskbarDarkColor, borderWidth, ButtonBorderStyle.Solid,
                    taskbarDarkColor, borderWidth, ButtonBorderStyle.Solid,
                    taskbarLightColor, borderWidth, ButtonBorderStyle.Solid,
                    taskbarLightColor, borderWidth, ButtonBorderStyle.Solid);
            }
            else
            {
                ControlPaint.DrawBorder(e.Graphics, rect,
                    taskbarLightColor, borderWidth, ButtonBorderStyle.Solid,
                    taskbarLightColor, borderWidth, ButtonBorderStyle.Solid,
                    taskbarDarkColor, borderWidth, ButtonBorderStyle.Solid,
                    taskbarDarkColor, borderWidth, ButtonBorderStyle.Solid);
            }
        }

        private bool IsButtonSelectedPressed(Button button)
        {
            if (button == btnStart)
            {
                return startMenu.Visible;
            }

            if (button.Tag is IntPtr handle && windowsByHandle.TryGetValue(handle, out var window))
            {
                return !window.IsMinimized && handle == lastActiveWindowHandle;
            }

            return false;
        }

        private void LoadQuickLaunchButtons()
        {
            if (InvokeRequired)
            {
                BeginInvoke(new Action(LoadQuickLaunchButtons));
                return;
            }

            foreach (var control in quickLaunchPanel.Controls.Cast<Control>().ToList())
            {
                if (control is Button quickLaunchButton)
                {
                    quickLaunchButton.Click -= QuickLaunchButton_Click;
                    quickLaunchButton.MouseUp -= QuickLaunchButton_MouseUp;
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
            using var icon = GetQuickLaunchButtonIcon(shortcutPath);
            var imageSize = GetScaledSmallIconSize();
            var image = icon != null ? ResizeIconImage(icon, imageSize) : ResizeIconImage(explorerMenuIcon, imageSize);
            var displayName = Path.GetFileNameWithoutExtension(shortcutPath);

            var button = new Button
            {
                Width = QuickLaunchButtonSize,
                Height = QuickLaunchButtonSize,
                TabStop = false,
                Margin = new Padding(0, 0, TaskbarButtonGap, 0),
                Padding = Padding.Empty,
                ImageAlign = ContentAlignment.MiddleCenter,
                Image = image,
                Tag = shortcutPath
            };
            ApplyQuickLaunchButtonStyle(button);
            AttachNoFocusCue(button);

            quickLaunchToolTip.SetToolTip(button, displayName);
            button.Click += QuickLaunchButton_Click;
            button.MouseUp += QuickLaunchButton_MouseUp;
            return button;
        }

        private static Image? GetQuickLaunchButtonIcon(string quickLaunchPath)
        {
            var extension = Path.GetExtension(quickLaunchPath);
            if (!extension.Equals(".lnk", StringComparison.OrdinalIgnoreCase))
            {
                return GetSmallShellIcon(quickLaunchPath, isDirectory: false);
            }

            var targetPath = ResolveShortcutTargetPath(quickLaunchPath);
            if (!string.IsNullOrWhiteSpace(targetPath) && (File.Exists(targetPath) || Directory.Exists(targetPath)))
            {
                var targetIcon = GetSmallShellIcon(targetPath, isDirectory: Directory.Exists(targetPath));
                if (targetIcon != null)
                {
                    return targetIcon;
                }
            }

            return GetSmallShellIcon(quickLaunchPath, isDirectory: false);
        }

        private static string? ResolveShortcutTargetPath(string shortcutPath)
        {
            object? shell = null;
            object? shortcut = null;

            try
            {
                var shellType = Type.GetTypeFromProgID("WScript.Shell");
                if (shellType == null)
                {
                    return null;
                }

                shell = Activator.CreateInstance(shellType);
                if (shell == null)
                {
                    return null;
                }

                shortcut = shellType.InvokeMember(
                    "CreateShortcut",
                    System.Reflection.BindingFlags.InvokeMethod,
                    binder: null,
                    target: shell,
                    args: new object[] { shortcutPath });

                if (shortcut == null)
                {
                    return null;
                }

                var shortcutType = shortcut.GetType();
                return shortcutType.InvokeMember(
                    "TargetPath",
                    System.Reflection.BindingFlags.GetProperty,
                    binder: null,
                    target: shortcut,
                    args: null) as string;
            }
            catch
            {
                return null;
            }
            finally
            {
                if (shortcut != null && Marshal.IsComObject(shortcut))
                {
                    Marshal.FinalReleaseComObject(shortcut);
                }

                if (shell != null && Marshal.IsComObject(shell))
                {
                    Marshal.FinalReleaseComObject(shell);
                }
            }
        }

        private Button CreateFallbackQuickLaunchButton(string text, Image icon, EventHandler onClick)
        {
            var button = new Button
            {
                Width = QuickLaunchButtonSize,
                Height = QuickLaunchButtonSize,
                TabStop = false,
                Margin = new Padding(0, 0, TaskbarButtonGap, 0),
                Padding = Padding.Empty,
                ImageAlign = ContentAlignment.MiddleCenter,
                Image = ResizeIconImage(icon, GetScaledSmallIconSize())
            };
            ApplyQuickLaunchButtonStyle(button);
            AttachNoFocusCue(button);

            quickLaunchToolTip.SetToolTip(button, text);
            button.Click += onClick;
            return button;
        }

        private void ResizeQuickLaunchPanel()
        {
            var totalWidth = quickLaunchPanel.Controls
                .OfType<Control>()
                .Sum(control => control.Width + control.Margin.Horizontal);

            quickLaunchPanel.Width = totalWidth > 0
                ? totalWidth + quickLaunchPanel.Padding.Horizontal
                : 0;
        }

        private void QuickLaunchButton_Click(object? sender, EventArgs e)
        {
            if (sender is not Button button || button.Tag is not string shortcutPath)
            {
                return;
            }

            LaunchProcess(shortcutPath);
        }

        private void ApplyQuickLaunchButtonStyle(Button button)
        {
            button.Paint -= TaskbarButtonWin98_Paint;
            button.Paint -= QuickLaunchButton_Paint;
            button.MouseDown -= QuickLaunchButton_MouseDownVisual;
            button.MouseUp -= QuickLaunchButton_MouseUpVisual;
            button.MouseLeave -= QuickLaunchButton_StateChangedVisual;
            button.LostFocus -= QuickLaunchButton_StateChangedVisual;
            button.MouseCaptureChanged -= QuickLaunchButton_StateChangedVisual;
            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderSize = 0;
            button.FlatAppearance.MouseOverBackColor = taskbarBaseColor;
            button.FlatAppearance.MouseDownBackColor = taskbarBaseColor;
            button.UseVisualStyleBackColor = false;
            button.BackColor = taskbarBaseColor;
            button.Paint += QuickLaunchButton_Paint;
            button.MouseDown += QuickLaunchButton_MouseDownVisual;
            button.MouseUp += QuickLaunchButton_MouseUpVisual;
            button.MouseLeave += QuickLaunchButton_StateChangedVisual;
            button.LostFocus += QuickLaunchButton_StateChangedVisual;
            button.MouseCaptureChanged += QuickLaunchButton_StateChangedVisual;
        }

        private void AttachNoFocusCue(Button button)
        {
            button.GotFocus += (_, _) =>
            {
                ActiveControl = null;
                BeginInvoke(new Action(() => button.Invalidate()));
            };
        }

        private static void QuickLaunchButton_MouseDownVisual(object? sender, MouseEventArgs e)
        {
            if (sender is Button button && e.Button == MouseButtons.Left)
            {
                button.Invalidate();
            }
        }

        private static void QuickLaunchButton_MouseUpVisual(object? sender, MouseEventArgs e)
        {
            if (sender is Button button)
            {
                button.Invalidate();
            }
        }

        private static void QuickLaunchButton_StateChangedVisual(object? sender, EventArgs e)
        {
            if (sender is Button button)
            {
                button.Invalidate();
            }
        }

        private void QuickLaunchButton_Paint(object? sender, PaintEventArgs e)
        {
            if (sender is not Button button)
            {
                return;
            }

            var isPressed = button.Capture && button.ClientRectangle.Contains(button.PointToClient(Cursor.Position));
            var rect = new Rectangle(0, 0, button.Width - 1, button.Height - 1);

            if (isPressed)
            {
                ControlPaint.DrawBorder(e.Graphics, rect,
                    taskbarDarkColor, 1, ButtonBorderStyle.Solid,
                    taskbarDarkColor, 1, ButtonBorderStyle.Solid,
                    taskbarLightColor, 1, ButtonBorderStyle.Solid,
                    taskbarLightColor, 1, ButtonBorderStyle.Solid);
            }
            else
            {
                ControlPaint.DrawBorder(e.Graphics, rect,
                    taskbarBaseColor, 1, ButtonBorderStyle.Solid,
                    taskbarBaseColor, 1, ButtonBorderStyle.Solid,
                    taskbarBaseColor, 1, ButtonBorderStyle.Solid,
                    taskbarBaseColor, 1, ButtonBorderStyle.Solid);
            }
        }

        private void QuickLaunchButton_MouseUp(object? sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Right || sender is not Button button || button.Tag is not string shortcutPath)
            {
                return;
            }

            quickLaunchContextShortcutPath = shortcutPath;
            quickLaunchItemContextMenu.Show(button, e.Location);
        }

        private void OpenQuickLaunchContextItem()
        {
            if (string.IsNullOrWhiteSpace(quickLaunchContextShortcutPath) || !File.Exists(quickLaunchContextShortcutPath))
            {
                return;
            }

            LaunchProcess(quickLaunchContextShortcutPath);
        }

        private void RemoveQuickLaunchContextItem()
        {
            if (string.IsNullOrWhiteSpace(quickLaunchContextShortcutPath) || !File.Exists(quickLaunchContextShortcutPath))
            {
                return;
            }

            try
            {
                File.Delete(quickLaunchContextShortcutPath);
                LoadQuickLaunchButtons();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to remove quick launch item: {ex.Message}");
            }
        }

        private void ShowQuickLaunchContextItemProperties()
        {
            if (string.IsNullOrWhiteSpace(quickLaunchContextShortcutPath) || !File.Exists(quickLaunchContextShortcutPath))
            {
                return;
            }

            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = quickLaunchContextShortcutPath,
                    Verb = "properties",
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to open quick launch item properties: {ex.Message}");
            }
        }

        private void QuickLaunchPanel_DragEnter(object? sender, DragEventArgs e)
        {
            e.Effect = e.Data?.GetDataPresent(DataFormats.FileDrop) == true
                ? DragDropEffects.Copy
                : DragDropEffects.None;
        }

        private void QuickLaunchPanel_DragDrop(object? sender, DragEventArgs e)
        {
            if (!e.Data!.GetDataPresent(DataFormats.FileDrop))
            {
                return;
            }

            var droppedPaths = e.Data.GetData(DataFormats.FileDrop) as string[];
            if (droppedPaths == null || droppedPaths.Length == 0)
            {
                return;
            }

            var quickLaunchFolder = GetQuickLaunchFolderPath();
            Directory.CreateDirectory(quickLaunchFolder);

            var createdAny = false;
            foreach (var path in droppedPaths)
            {
                try
                {
                    if (!File.Exists(path) && !Directory.Exists(path))
                    {
                        continue;
                    }

                    var extension = Path.GetExtension(path);
                    if (extension.Equals(".lnk", StringComparison.OrdinalIgnoreCase) ||
                        extension.Equals(".url", StringComparison.OrdinalIgnoreCase) ||
                        extension.Equals(".scf", StringComparison.OrdinalIgnoreCase))
                    {
                        var destination = GetUniqueFilePath(quickLaunchFolder, Path.GetFileName(path));
                        File.Copy(path, destination, overwrite: false);
                        createdAny = true;
                        continue;
                    }

                    var displayName = Path.GetFileNameWithoutExtension(path);
                    var shortcutPath = GetUniqueFilePath(quickLaunchFolder, displayName + ".lnk");
                    CreateShellShortcut(shortcutPath, path);
                    createdAny = true;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Failed to add Quick Launch item '{path}': {ex.Message}");
                }
            }

            if (createdAny)
            {
                LoadQuickLaunchButtons();
            }
        }

        private static string GetUniqueFilePath(string folderPath, string fileName)
        {
            var destination = Path.Combine(folderPath, fileName);
            if (!File.Exists(destination))
            {
                return destination;
            }

            var baseName = Path.GetFileNameWithoutExtension(fileName);
            var extension = Path.GetExtension(fileName);
            var index = 2;
            while (true)
            {
                var candidate = Path.Combine(folderPath, $"{baseName} ({index}){extension}");
                if (!File.Exists(candidate))
                {
                    return candidate;
                }

                index++;
            }
        }

        private static void CreateShellShortcut(string shortcutPath, string targetPath)
        {
            var shellType = Type.GetTypeFromProgID("WScript.Shell");
            if (shellType == null)
            {
                return;
            }

            dynamic shell = Activator.CreateInstance(shellType)!;
            dynamic shortcut = shell.CreateShortcut(shortcutPath);
            shortcut.TargetPath = targetPath;
            shortcut.WorkingDirectory = Directory.Exists(targetPath)
                ? targetPath
                : (Path.GetDirectoryName(targetPath) ?? string.Empty);
            shortcut.Save();
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

        private static void OpenTaskbarSettings()
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "ms-settings:taskbar",
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
                    Arguments = "/name Microsoft.TaskbarAndStartMenu",
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to open taskbar settings: {ex.Message}");
            }
        }

        private static void OpenFolderOptionsSettings()
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "control.exe",
                    Arguments = "folders",
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to open folder options: {ex.Message}");
            }
        }

        private static void OpenNetworkConnectionsSettings()
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "control.exe",
                    Arguments = "ncpa.cpl",
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
                    FileName = "ms-settings:network",
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to open network settings: {ex.Message}");
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

        private void ResizeGripPanel_Paint(object? sender, PaintEventArgs e)
        {
            using var lightPen = new Pen(taskbarLightColor);
            var lines = Math.Max(1, ResizeGripHeight);

            if (taskbarLocked)
            {
                for (var i = 0; i < lines && i < resizeGripPanel.Height; i++)
                {
                    e.Graphics.DrawLine(lightPen, 0, i, resizeGripPanel.Width, i);
                }

                return;
            }

            using var darkPen = new Pen(taskbarDarkColor);
            var lightLines = Math.Max(1, lines - 1);
            for (var i = 0; i < lightLines && i < resizeGripPanel.Height; i++)
            {
                e.Graphics.DrawLine(lightPen, 0, i, resizeGripPanel.Width, i);
            }

            for (var i = lightLines; i < lines; i++)
            {
                if (i >= resizeGripPanel.Height)
                {
                    break;
                }

                e.Graphics.DrawLine(darkPen, 0, i, resizeGripPanel.Width, i);
            }
        }

        private void NotificationAreaPanel_Paint(object? sender, PaintEventArgs e)
        {
            var rect = new Rectangle(0, 0, notificationAreaPanel.Width - 1, notificationAreaPanel.Height - 1);
            if (rect.Width <= 1 || rect.Height <= 1)
            {
                return;
            }

            using var darkPen = new Pen(taskbarDarkColor);
            using var lightPen = new Pen(taskbarLightColor);

            // Recessed 1px border: top/left dark, bottom/right light.
            e.Graphics.DrawLine(darkPen, rect.Left, rect.Top, rect.Right, rect.Top);
            e.Graphics.DrawLine(darkPen, rect.Left, rect.Top, rect.Left, rect.Bottom);
            e.Graphics.DrawLine(lightPen, rect.Right, rect.Top, rect.Right, rect.Bottom);
            e.Graphics.DrawLine(lightPen, rect.Left, rect.Bottom, rect.Right, rect.Bottom);
        }

        private void UpdateResizeUiState()
        {
            resizeGripPanel.Visible = true;
            resizeGripPanel.Height = ResizeGripHeight;
            resizeGripPanel.Cursor = taskbarLocked ? Cursors.Default : Cursors.SizeNS;
            resizeGripPanel.Invalidate();
        }

        private void SetTaskbarLocked(bool locked)
        {
            taskbarLocked = locked;
            UpdateResizeUiState();
            SetTaskbarBounds();
            RefreshTaskbar();
            SaveTaskbarSettings();
        }

        private void TaskbarResize_MouseDown(object? sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left || taskbarLocked)
            {
                return;
            }

            var resizeHotZone = ResizeGripHeight + 1;
            if (e.Y > resizeHotZone)
            {
                return;
            }

            isTaskbarResizeDragging = true;
            Capture = true;
        }

        private void TaskbarResize_MouseMove(object? sender, MouseEventArgs e)
        {
            if (!isTaskbarResizeDragging)
            {
                return;
            }

            var screenBounds = Screen.PrimaryScreen?.Bounds ?? new Rectangle(0, 0, 1024, 768);
            var newHeight = screenBounds.Bottom - Cursor.Position.Y;
            var newRows = Math.Clamp((int)Math.Round(newHeight / (double)TaskbarRowHeight), 1, 3);

            if (newRows != taskbarRows)
            {
                taskbarRows = newRows;
                SetTaskbarBounds();
                RefreshTaskbar();
            }
        }

        private void TaskbarResize_MouseUp(object? sender, MouseEventArgs e)
        {
            if (!isTaskbarResizeDragging)
            {
                return;
            }

            isTaskbarResizeDragging = false;
            Capture = false;
            SaveTaskbarSettings();
        }

        private void ChildTaskbarResize_MouseDown(object? sender, MouseEventArgs e)
        {
            if (sender is not Control child)
            {
                return;
            }

            var local = PointToClient(child.PointToScreen(e.Location));
            TaskbarResize_MouseDown(this, new MouseEventArgs(e.Button, e.Clicks, local.X, local.Y, e.Delta));
        }

        private void ChildTaskbarResize_MouseMove(object? sender, MouseEventArgs e)
        {
            if (!isTaskbarResizeDragging)
            {
                return;
            }

            TaskbarResize_MouseMove(this, e);
        }

        private void ChildTaskbarResize_MouseUp(object? sender, MouseEventArgs e)
        {
            if (!isTaskbarResizeDragging)
            {
                return;
            }

            TaskbarResize_MouseUp(this, e);
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
                playVolumeFeedbackSound,
                startMenuSubmenuOpenDelayMs,
                autoHideTaskbar,
                taskbarButtonStyle.ToString(),
                taskbarBevelSize,
                taskbarFontName,
                taskbarFontSize,
                taskbarBaseColor,
                taskbarLightColor,
                taskbarDarkColor,
                OpenQuickLaunchFolderInExplorerForm,
                () => Application.Exit());
            dialog.ApplyRequested += () => ApplyOptionsFromDialog(dialog);
            if (dialog.ShowDialog(this) != DialogResult.OK)
            {
                return;
            }

            ApplyOptionsFromDialog(dialog);
        }

        private void ApplyOptionsFromDialog(TaskbarOptionsForm dialog)
        {
            startMenuIconSize = dialog.StartMenuIconSize;
            taskIconSize = dialog.TaskIconSize;
            lazyLoadProgramsSubmenu = dialog.LazyLoadProgramsSubmenu;
            playVolumeFeedbackSound = dialog.PlayVolumeFeedbackSound;
            startMenuSubmenuOpenDelayMs = dialog.StartMenuSubmenuOpenDelayMs;
            autoHideTaskbar = dialog.AutoHideTaskbar;
            taskbarBevelSize = Math.Clamp(dialog.TaskbarBevelSize, 1, 4);
            taskbarFontName = dialog.TaskbarFontName;
            taskbarFontSize = dialog.TaskbarFontSize;
            taskbarBaseColor = dialog.TaskbarBaseColor;
            taskbarLightColor = dialog.TaskbarLightColor;
            taskbarDarkColor = dialog.TaskbarDarkColor;
            if (Enum.TryParse<TaskbarButtonStyle>(dialog.TaskbarButtonStyle, ignoreCase: true, out var parsedStyle))
            {
                taskbarButtonStyle = parsedStyle;
            }
            volumePopup.PlayFeedbackSoundOnMouseUp = playVolumeFeedbackSound;
            ApplyIconSizeSettings();
            SaveTaskbarSettings();
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
            RefreshTaskbarFonts();
            RefreshTaskbarColors();

            var oldStartImage = btnStart.Image;
            btnStart.Image = AddIconMargin(ResizeIconImage(startButtonMenuIcon, startMenuIconSize), 2, 1);
            oldStartImage?.Dispose();
            ApplyTaskbarButtonStyle(btnStart);

            startMenu.Dispose();
            startMenu = BuildStartMenu();
            InvalidateProgramMenuCache();

            LoadQuickLaunchButtons();

            foreach (var button in taskButtons.Values)
            {
                var oldImage = button.Image;
                button.Image = null;
                oldImage?.Dispose();
                ApplyTaskbarButtonStyle(button);
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
            if ((DateTime.UtcNow - lastStartMenuClosedAtUtc).TotalMilliseconds < 250)
            {
                startMenuWasVisibleOnStartMouseDown = false;
                return;
            }

            if (startMenuWasVisibleOnStartMouseDown)
            {
                if (startMenu.Visible)
                {
                    startMenu.Hide();
                }

                ResetStartMenuSearch();
                ResetStartMenuSubmenuSessionCache();
                btnStart.Invalidate();
                startMenuWasVisibleOnStartMouseDown = false;
                return;
            }

            ToggleStartMenu();
            startMenuWasVisibleOnStartMouseDown = false;
        }

        private void BtnStart_MouseDown(object? sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left)
            {
                return;
            }

            startMenuWasVisibleOnStartMouseDown = startMenu.Visible;
        }

        private void ToggleStartMenu()
        {
            if (startMenu.Visible)
            {
                startMenu.Hide();
                ResetStartMenuSearch();
                ResetStartMenuSubmenuSessionCache();
                btnStart.Invalidate();
                return;
            }

            ResetStartMenuSearch();
            ResetStartMenuSubmenuSessionCache();

            startMenu.Show(btnStart, new Point(0, -Math.Max(startMenu.GetPreferredSize(Size.Empty).Height, 1)));
            RepositionVisibleStartMenu();
            var firstItem = startMenu.Items.OfType<ToolStripItem>().FirstOrDefault(IsNavigableTopLevelMenuItem);
            firstItem?.Select();
            btnStart.Invalidate();
        }

        private void RepositionVisibleStartMenu()
        {
            if (!startMenu.Visible)
            {
                return;
            }

            var startButtonLocation = btnStart.PointToScreen(Point.Empty);
            var menuX = startButtonLocation.X;
            var menuHeight = Math.Max(startMenu.GetPreferredSize(Size.Empty).Height, 1);
            var menuY = startButtonLocation.Y - menuHeight;
            var screenBounds = Screen.FromPoint(startButtonLocation).WorkingArea;

            if (menuX + startMenu.Width > screenBounds.Right)
            {
                menuX = screenBounds.Right - startMenu.Width;
            }

            if (menuY < screenBounds.Top)
            {
                menuY = screenBounds.Top;
            }

            startMenu.Location = new Point(menuX, menuY);
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
                trayToolTip.SetToolTip(control, GetTrayHoverMessage(icon));
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
            control.MouseDoubleClick += TrayIcon_MouseDoubleClick;
            control.MouseMove += TrayIcon_MouseMove;
            control.MouseEnter += TrayIcon_MouseEnter;
            control.MouseLeave += TrayIcon_MouseLeave;

            UpdateTrayIconImage(control, icon);
            return control;
        }

        private void UpdateTrayIconImage(PictureBox control, TrayNotifyIcon icon)
        {
            var oldImage = control.Image;
            var image = ConvertImageSourceToBitmap(icon.Icon, GetScaledSmallIconSize());
            control.Image = image;
            oldImage?.Dispose();
        }

        private void ResizeTrayPanel()
        {
            var count = trayIconControls.Count + 1;
            trayIconsPanel.Width = Math.Clamp((count * 20) + 8, 32, 220);
            notificationAreaPanel.Width = trayIconsPanel.Width + lblClock.Width + notificationAreaPanel.Padding.Horizontal + 6;
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
            control.MouseDoubleClick -= TrayIcon_MouseDoubleClick;
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

        private void TrayIcon_MouseDoubleClick(object? sender, MouseEventArgs e)
        {
            if (sender is not PictureBox control || control.Tag is not TrayNotifyIcon icon)
            {
                return;
            }

            if (e.Button == MouseButtons.Left && IsWindowsSecurityTrayIcon(icon))
            {
                OpenWindowsSecurityCenter();
                return;
            }
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

        private static bool IsWindowsSecurityTrayIcon(TrayNotifyIcon icon)
        {
            if (!string.IsNullOrWhiteSpace(icon.Path) &&
                icon.Path.Contains("SecurityHealthSystray", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (!string.IsNullOrWhiteSpace(icon.Identifier) &&
                icon.Identifier.Contains("SecurityHealth", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (!string.IsNullOrWhiteSpace(icon.Title) &&
                (icon.Title.Contains("Windows Security", StringComparison.OrdinalIgnoreCase) ||
                 icon.Title.Contains("Windows 安全", StringComparison.OrdinalIgnoreCase) ||
                 icon.Title.Contains("安全性", StringComparison.OrdinalIgnoreCase)))
            {
                return true;
            }

            return false;
        }

        private static string GetTrayHoverMessage(TrayNotifyIcon icon)
        {
            if (!string.IsNullOrWhiteSpace(icon.Title))
            {
                return icon.Title.Trim();
            }

            if (!string.IsNullOrWhiteSpace(icon.Identifier))
            {
                return icon.Identifier.Trim();
            }

            if (!string.IsNullOrWhiteSpace(icon.Path))
            {
                return Path.GetFileNameWithoutExtension(icon.Path);
            }

            return "Tray icon";
        }

        private static void OpenWindowsSecurityCenter()
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "windowsdefender:",
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
                    FileName = "ms-settings:windowsdefender",
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to open Windows Security: {ex.Message}");
            }
        }

        private void ShowVolumePopupNear(Control trayIconControl)
        {
            if (IsDisposed || Disposing || volumePopup.IsDisposed)
            {
                return;
            }

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
                        // Handle Win+D for Show Desktop
                        if (vkCode == 0x44) // D key
                        {
                            BeginInvoke(new Action(ShowDesktop));
                            nonWinKeyPressedWhileWinHeld = true;
                            return (IntPtr)1;
                        }
                        // Let other Win+ combinations pass through to the system
                        nonWinKeyPressedWhileWinHeld = true;
                        return (IntPtr)0; // Don't block other Win+ shortcuts
                    }

                    if (HandleStartMenuSearchKey(vkCode))
                    {
                        return (IntPtr)1;
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
            else
            {
                var activeByState = windows
                    .FirstOrDefault(w => w.State == ApplicationWindow.WindowState.Active)?.Handle ?? IntPtr.Zero;

                if (currentHandles.Contains(activeByState))
                {
                    lastActiveWindowHandle = activeByState;
                }
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
                TabStop = false,
                Margin = new Padding(0, 0, TaskbarButtonGap, 0),
                TextAlign = ContentAlignment.MiddleLeft,
                TextImageRelation = TextImageRelation.ImageBeforeText,
                ImageAlign = ContentAlignment.MiddleLeft,
                Font = normalFont
            };
            ApplyTaskbarButtonStyle(button);
            AttachNoFocusCue(button);

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
                    button.Image = AddIconMargin(image, 2, 1);
                }
            }

            var isActive = window.Handle == lastActiveWindowHandle || window.State == ApplicationWindow.WindowState.Active;

            if (isActive)
            {
                button.Font = activeFont;
                button.BackColor = taskbarLightColor;
            }
            else
            {
                button.Font = normalFont;
                button.BackColor = taskbarBaseColor;
            }

            if (window.State == ApplicationWindow.WindowState.Flashing)
            {
                button.BackColor = Color.FromArgb(255, 255, 192);
            }

            button.Text = TruncateTaskButtonText(button, title, button.Font);

            if (taskbarButtonStyle == TaskbarButtonStyle.Win98)
            {
                button.Invalidate();
            }
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

        private static Bitmap AddIconMargin(Image source, int rightMargin, int bottomMargin)
        {
            var bitmap = new Bitmap(source.Width + rightMargin, source.Height + bottomMargin, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using var g = Graphics.FromImage(bitmap);
            g.Clear(Color.Transparent);
            g.DrawImage(source, 0, 0, source.Width, source.Height);
            source.Dispose();
            return bitmap;
        }

        private int GetScaledSmallIconSize()
        {
            var scale = DeviceDpi <= 0 ? 1.0 : DeviceDpi / 96.0;
            return Math.Clamp((int)Math.Round(16 * scale), 16, 18);
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
                if (window.IsMinimized || WinApi.IsIconic(handle))
                {
                    WinApi.ShowWindow(handle, WinApi.SW_RESTORE);
                }

                window.BringToFront();
                WinApi.SetForegroundWindow(handle);
                lastActiveWindowHandle = handle;
            }
            else if (handle == foregroundWindowBeforeTaskClick || handle == lastActiveWindowHandle)
            {
                window.Minimize();
                if (!window.IsMinimized && !WinApi.IsIconic(handle))
                {
                    WinApi.ShowWindow(handle, WinApi.SW_MINIMIZE);
                }

                if (lastActiveWindowHandle == handle)
                {
                    lastActiveWindowHandle = IntPtr.Zero;
                }
            }
            else
            {
                window.BringToFront();
                WinApi.SetForegroundWindow(handle);
                lastActiveWindowHandle = handle;
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
            var totalHeight = TotalTaskbarHeight;
            var taskbarY = screenBounds.Bottom - totalHeight;

            if (autoHideTaskbar)
            {
                var cursor = Cursor.Position;
                var nearBottomEdge = cursor.X >= screenBounds.Left &&
                    cursor.X <= screenBounds.Right &&
                    cursor.Y >= screenBounds.Bottom - 2;
                var cursorOverExpandedArea = cursor.X >= screenBounds.Left &&
                    cursor.X <= screenBounds.Right &&
                    cursor.Y >= screenBounds.Bottom - totalHeight;
                var keepVisible = nearBottomEdge || cursorOverExpandedArea ||
                    startMenu.Visible || taskWindowMenu.Visible || taskbarContextMenu.Visible ||
                    clockContextMenu.Visible || quickLaunchItemContextMenu.Visible ||
                    volumePopup.Visible || timeDetailsPopup.Visible || ContainsFocus;

                taskbarY = keepVisible ? screenBounds.Bottom - totalHeight : screenBounds.Bottom - 2;
            }

            var newBounds = new Rectangle(screenBounds.Left, taskbarY, screenBounds.Width, totalHeight);

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
