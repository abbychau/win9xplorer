using System.Diagnostics;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Drawing.Drawing2D;
using System.Globalization;
using Microsoft.Win32;
using Microsoft.VisualBasic;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Media.Imaging;
using ManagedShell;
using ManagedShell.WindowsTray;
using ManagedShell.WindowsTasks;
using MediaImageSource = System.Windows.Media.ImageSource;
using TrayNotifyIcon = ManagedShell.WindowsTray.NotifyIcon;
using WpfMouseButton = System.Windows.Input.MouseButton;

namespace win9xplorer
{
    internal sealed partial class RetroTaskbarForm : Form
    {
        private const int CLSCTX_ALL = 23;
        private static readonly Guid CLSID_MMDeviceEnumerator = new("BCDE0395-E52F-467C-8E3D-C4579291692E");
        private static readonly Guid IID_IAudioEndpointVolume = new("5CDF2C82-841E-4546-9722-0CF74078229A");

        private static RetroTaskbarForm? instance;

        private enum EDataFlow
        {
            eRender,
            eCapture,
            eAll,
            EDataFlow_enum_count
        }

        private enum ERole
        {
            eConsole,
            eMultimedia,
            eCommunications,
            ERole_enum_count
        }

        [ComImport]
        [Guid("A95664D2-9614-4F35-A746-DE8DB63617E6")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IMMDeviceEnumerator
        {
            int NotImpl1();
            int GetDefaultAudioEndpoint(EDataFlow dataFlow, ERole role, out IMMDevice device);
        }

        [ComImport]
        [Guid("D666063F-1587-4E43-81F1-B948E807363F")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IMMDevice
        {
            int Activate(ref Guid iid, int clsCtx, IntPtr activationParams, [MarshalAs(UnmanagedType.Interface)] out object interfacePointer);
        }

        [ComImport]
        [Guid("5CDF2C82-841E-4546-9722-0CF74078229A")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IAudioEndpointVolume
        {
            int RegisterControlChangeNotify(IntPtr notify);
            int UnregisterControlChangeNotify(IntPtr notify);
            int GetChannelCount(out uint channelCount);
            int SetMasterVolumeLevel(float levelDb, Guid eventContext);
            int SetMasterVolumeLevelScalar(float level, Guid eventContext);
            int GetMasterVolumeLevel(out float levelDb);
            int GetMasterVolumeLevelScalar(out float level);
        }
        private const int TaskbarRowHeight = 30;
        private const int TaskbarButtonVerticalPadding = 2;
        private const int TaskbarBottomPadding = 0;
        private const int TaskbarButtonGap = 2;
        private const int AddressRowSeparatorHeight = 2;
        private const int ResizeGripMinHeight = 1;
        private const int RefreshIntervalMs = 120;
        private const int DefaultStartMenuIconSize = 16;
        private const int DefaultTaskIconSize = 16;
        private const int TaskButtonMaxWidth = 220;
        private const int TaskButtonMinWidthMultiRow = 120;
        private const int TaskButtonRowBottomGap = 1;
        private const int ProgramMenuMaxDepth = 4;
        private const int StartMenuSearchResetSeconds = 2;
        private const int MultiColumnThreshold = 22;
        private const int MaxSubmenuColumns = 3;
        private const int TrayIconCellSize = 20;
        private const int TrayIconNewIconStabilityMs = 650;
        private const int AddressHistoryMaxEntries = 40;
        private const int DefaultStartMenuRecentProgramsMaxCount = 10;
        private const string StartMenuCommandTokenPrefix = "command::";
        private const int ClockLineSpacingPx = 2;
        private static readonly Color Win98ActiveTaskButtonBackColor = Color.FromArgb(224, 224, 224);
        private const int WS_EX_TOOLWINDOW = 0x00000080;
        private const int WS_EX_APPWINDOW = 0x00040000;
        private static readonly Color InputIndicatorBackColor = Color.FromArgb(0, 0, 128);
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
        private readonly Panel taskbarResizeHandlePanel;
        private readonly Panel contentPanel;
        private readonly Panel startHostPanel;
        private readonly Panel taskButtonsHostPanel;
        private readonly Panel addressRowHostPanel;
        private readonly Panel addressToolbarPanel;
        private readonly Panel volumeToolbarPanel;
        private readonly Panel spotifyToolbarPanel;
        private readonly Panel spotifyControlsPanel;
        private readonly Panel startToolbarSeparatorPanel;
        private readonly Panel toolbarsSeparatorPanel;
        private readonly Panel spotifyToolbarSeparatorPanel;
        private readonly Panel taskButtonsSeparatorPanel;
        private readonly Panel addressToolbarSeparatorPanel;
        private readonly Panel addressRowSeparatorPanel;
        private readonly GripPanel quickLaunchGripPanel;
        private readonly GripPanel addressToolbarGripPanel;
        private readonly GripPanel addressRowGripPanel;
        private readonly GripPanel volumeToolbarGripPanel;
        private readonly GripPanel spotifyToolbarGripPanel;
        private readonly GripPanel runningProgramsGripPanel;
        private readonly List<ToolbarGripRelation> toolbarGripRelations = new();
        private readonly List<ToolbarSeparatorRelation> toolbarSeparatorRelations = new();
        private readonly Panel addressInputHostPanel;
        private readonly Panel notificationAreaPanel;
        private readonly Button btnStart;
        private readonly Button btnSpotifyPrevious;
        private readonly Button btnSpotifyPlayPause;
        private readonly Button btnSpotifyNext;
        private readonly Label spotifyNowPlayingLabel;
        private readonly ComboBox addressToolbarComboBox;
        private readonly TrackBar volumeToolbarTrackBar;
        private IntPtr appBarHandle = IntPtr.Zero;
        private bool isAppBarRegistered = false;
        private bool isFullscreenActive = false;
        private bool suppressFullscreenAutoDetect;
        private readonly FlowLayoutPanel quickLaunchPanel;
        private readonly FlowLayoutPanel taskButtonsPanel;
        private readonly FlowLayoutPanel trayIconsPanel;
        private readonly Label lblClock;
        private readonly Label languageIndicatorLabel;
        private readonly Label imeIndicatorLabel;
        private readonly System.Windows.Forms.Timer refreshTimer;
        private readonly System.Windows.Forms.Timer trayResizeShrinkTimer;
        private readonly System.Windows.Forms.Timer trayIconStabilityTimer;
        private ContextMenuStrip startMenu;
        private readonly ContextMenuStrip taskWindowMenu;
        private readonly ContextMenuStrip taskbarContextMenu;
        private readonly ContextMenuStrip clockContextMenu;
        private readonly ContextMenuStrip languageIndicatorContextMenu;
        private readonly ContextMenuStrip imeIndicatorContextMenu;
        private readonly ContextMenuStrip quickLaunchItemContextMenu;
        private readonly ToolTip taskToolTip;
        private readonly ToolTip trayToolTip;
        private readonly ToolTip quickLaunchToolTip;
        private ToolStripRenderer win9xMenuRenderer;
        private Font normalFont;
        private Font activeFont;
        private static readonly Font InputIndicatorFont = new Font("\u65B0\u7D30\u660E\u9AD4", 9.0f, FontStyle.Regular, GraphicsUnit.Point);
        private readonly Image folderMenuIcon;
        private readonly Image startButtonMenuIcon;
        private readonly Image programsMenuIcon;
        private readonly Image explorerMenuIcon;
        private readonly Image documentsMenuIcon;
        private readonly Image settingsMenuIcon;
        private readonly Image runMenuIcon;
        private readonly Image quickLaunchDesktopIcon;
        private readonly Image customVolumeIconImage;
        private readonly Image spotifyPreviousIconImage;
        private readonly Image spotifyPlayIconImage;
        private readonly Image spotifyNextIconImage;
        private sealed record ProgramMenuNode(string FolderPath, int Depth);
        private readonly Dictionary<IntPtr, Button> taskButtons = new();
        private readonly Dictionary<IntPtr, ApplicationWindow> windowsByHandle = new();
        private readonly HashSet<IntPtr> windowsMinimizedByShowDesktop = new();
        private readonly Dictionary<TrayNotifyIcon, PictureBox> trayIconControls = new();
        private readonly Dictionary<TrayNotifyIcon, DateTime> trayIconFirstSeenUtc = new();
        private int trayIconSlotCount;
        private bool trayInitialSyncComplete;
        private IntPtr contextMenuWindowHandle = IntPtr.Zero;
        private readonly TaskbarInteractionStateMachine interactionStateMachine;
        private readonly StartMenuController startMenuController;
        private readonly QuickLaunchController quickLaunchController = new();
        private readonly SpotifyController spotifyController = new();
        private readonly TrayController trayController = new();
        private readonly ThemeController themeController = new();
        private DiagnosticsPanelForm? diagnosticsPanel;
        private TaskManagerForm? taskManagerForm;
        private bool allowAppExit;
        private readonly WinApi.LowLevelKeyboardProc keyboardProc;
        private IntPtr keyboardHookHandle = IntPtr.Zero;
        private readonly WinApi.LowLevelMouseProc mouseProc;
        private IntPtr mouseHookHandle = IntPtr.Zero;
        private bool winKeyPressed;
        private bool nonWinKeyPressedWhileWinHeld;
        private bool injectedWinDownForChord;
        private int activeWinKeyVkCode;
        private readonly VolumePopupForm volumePopup;
        private TimeDetailsForm? timeDetailsPopup;
        private readonly PictureBox customVolumeIcon;
        private readonly TrayIconInteractionService trayIconInteractionService = new();
        private readonly QuickLaunchRefreshMonitor quickLaunchRefreshMonitor;
        private string? quickLaunchContextShortcutPath;
        private string? quickLaunchDragSourcePath;
        private Point quickLaunchDragStartPoint;
        private int quickLaunchPreviewInsertPosition = -1;
        private string selectedThemeProfileName = "Custom";
        private int startMenuIconSize = DefaultStartMenuIconSize;
        private int taskIconSize = DefaultTaskIconSize;
        private string taskbarFontName = "\u65B0\u7D30\u660E\u9AD4";
        private float taskbarFontSize = 9.0f;
        private Color taskbarFontColor = Color.Black;
        private bool lazyLoadProgramsSubmenu = true;
        private bool playVolumeFeedbackSound = true;
        private bool useClassicVolumePopup;
        private int startMenuSubmenuOpenDelayMs = 200;
        private int startMenuRecentProgramsMaxCount = DefaultStartMenuRecentProgramsMaxCount;
        private bool autoHideTaskbar;
        private bool startOnWindowsStartup;
        private bool showQuickLaunchToolbar = true;
        private bool showAddressToolbar;
        private bool showSpotifyToolbar = true;
        private bool showVolumeToolbar = true;
        private bool isSpotifyRunning;
        private DateTime spotifyStateLastCheckedUtc = DateTime.MinValue;
        private bool addressToolbarBeforeQuickLaunch;
        private string lastAddressToolbarText = string.Empty;
        private bool isToolbarReorderDragging;
        private bool toolbarOrderChangedDuringDrag;
        private string? draggingToolbarId;
        private bool suppressVolumeToolbarEvents;
        private bool isVolumeToolbarDragging;
        private TaskbarButtonStyle taskbarButtonStyle = TaskbarButtonStyle.Modern;
        private int taskbarBevelSize = 1;
        private Color taskbarBaseColor = Color.FromArgb(192, 192, 192);
        private Color taskbarLightColor = Color.FromArgb(255, 255, 255);
        private Color taskbarDarkColor = Color.FromArgb(128, 128, 128);
        private bool taskbarLocked;
        private int taskbarRows = 1;
        private bool isTaskbarResizeDragging;
        private readonly Dictionary<string, ProgramFolderCacheEntry> programFolderCache = new(StringComparer.OrdinalIgnoreCase);
        private sealed record ProgramFolderCacheEntry(List<string> Directories, List<string> Files);
        private readonly object shellIconCacheLock = new();
        private readonly Dictionary<string, Bitmap> shellIconCache = new(StringComparer.OrdinalIgnoreCase);
        private readonly StringBuilder startMenuSearchQuery = new();
        private DateTime startMenuSearchLastInputUtc = DateTime.MinValue;
        private readonly List<ToolStripItem> startMenuSearchResultItems = new();
        private readonly List<ToolStripItem> startMenuRecentProgramItems = new();
        private ToolStripTextBox? startMenuSearchTextBox;
        private bool suppressStartMenuSearchTextChanged;
        private readonly List<string> addressToolbarHistory = new();
        private readonly List<ProgramSearchEntry> programSearchEntriesCache = new();
        private readonly object programSearchEntriesCacheLock = new();
        private readonly List<ProgramSearchEntry> startMenuRecentPrograms = new();
        private readonly Dictionary<ToolStripMenuItem, EventHandler> menuDropDownOpeningHandlers = new();
        private readonly ToolbarComponents quickLaunchToolbar;
        private readonly ToolbarComponents addressToolbar;
        private readonly ToolbarComponents volumeToolbar;
        private readonly ToolbarComponents spotifyToolbar;
        private readonly IReadOnlyDictionary<ToolbarId, ToolbarComponents> toolbarsById;
        private ToolbarHostApplier? toolbarHostApplier;
        private string clockLineDate = string.Empty;
        private string clockLineWeekday = string.Empty;
        private string clockLineTime = string.Empty;
        private readonly HashSet<ToolStripMenuItem> populatedSubmenusThisSession = new();
        private sealed record ProgramSearchEntry(string DisplayName, string LaunchPath, bool IsShellApp);
        private sealed class TrayIconInteractionService
        {
            public void ForwardMouseDown(TrayNotifyIcon icon, MouseButtons button, uint cursorPosition, int doubleClickTime)
            {
                icon.IconMouseDown(ConvertMouseButton(button), cursorPosition, doubleClickTime);
            }

            public void ForwardMouseUp(TrayNotifyIcon icon, MouseButtons button, uint cursorPosition, int doubleClickTime)
            {
                icon.IconMouseUp(ConvertMouseButton(button), cursorPosition, doubleClickTime);
            }

            public void ForwardMouseMove(TrayNotifyIcon icon, uint cursorPosition)
            {
                icon.IconMouseMove(cursorPosition);
            }

            public void ForwardMouseEnter(TrayNotifyIcon icon, uint cursorPosition)
            {
                icon.IconMouseEnter(cursorPosition);
            }

            public void ForwardMouseLeave(TrayNotifyIcon icon, uint cursorPosition)
            {
                icon.IconMouseLeave(cursorPosition);
            }
        }

        private enum TaskbarButtonStyle
        {
            Modern,
            Win98
        }

        private sealed class GripPanel : Panel
        {
            private readonly RetroTaskbarForm owner;
            private readonly Cursor toolbarCursor;
            private readonly bool enableToolbarReorder;
            private readonly string? dragTag;

            public GripPanel(
                RetroTaskbarForm owner,
                Color backColor,
                Cursor cursor,
                bool visible = false,
                bool enabled = true,
                bool enableToolbarReorder = false,
                string? dragTag = null,
                bool forwardChildResizeMouse = false)
            {
                this.owner = owner;
                toolbarCursor = cursor;
                this.enableToolbarReorder = enableToolbarReorder;
                this.dragTag = dragTag;
                Dock = DockStyle.Left;
                Width = 8;
                Margin = Padding.Empty;
                Padding = Padding.Empty;
                BackColor = backColor;
                Cursor = cursor;
                Visible = visible;
                Enabled = enabled;
                Tag = dragTag;
                if (forwardChildResizeMouse)
                {
                    MouseDown += owner.ChildTaskbarResize_MouseDown;
                    MouseMove += owner.ChildTaskbarResize_MouseMove;
                    MouseUp += owner.ChildTaskbarResize_MouseUp;
                }
            }

            public void ApplyTheme(Color backColor)
            {
                BackColor = backColor;
            }

            public void SetToolbarVisible(bool visible)
            {
                Visible = visible;
            }

            public void ApplyCursor(bool locked)
            {
                Cursor = locked ? Cursors.Default : toolbarCursor;
            }

            public void RefreshGrip()
            {
                Invalidate();
            }

            public void ApplyLockState(bool locked)
            {
                ApplyCursor(locked);
                RefreshGrip();
            }

            public static void ApplyLockState(bool locked, params GripPanel[] grips)
            {
                foreach (var grip in grips)
                {
                    grip.ApplyLockState(locked);
                }
            }

            public static void ApplyTheme(Color backColor, params GripPanel[] grips)
            {
                foreach (var grip in grips)
                {
                    grip.ApplyTheme(backColor);
                }
            }

            public static void SetToolbarVisible(bool visible, params GripPanel[] grips)
            {
                foreach (var grip in grips)
                {
                    grip.SetToolbarVisible(visible);
                }
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                e.Graphics.Clear(owner.taskbarBaseColor);
                var left = 2;
                var top = 2;
                var bottom = Math.Max(top, Height - 5);
                owner.DrawToolbarGrip(e.Graphics, left, top, bottom);

            }

            protected override void OnMouseDown(MouseEventArgs e)
            {
                if (enableToolbarReorder)
                {
                    owner.ToolbarGripPanel_MouseDown(this, e);
                }

                base.OnMouseDown(e);
            }

            protected override void OnMouseMove(MouseEventArgs e)
            {
                if (enableToolbarReorder)
                {
                    owner.ToolbarGripPanel_MouseMove(this, e);
                }

                base.OnMouseMove(e);
            }

            protected override void OnMouseUp(MouseEventArgs e)
            {
                if (enableToolbarReorder)
                {
                    owner.ToolbarGripPanel_MouseUp(this, e);
                }

                base.OnMouseUp(e);
            }
        }

        private sealed record OptionsState(
            int StartMenuIconSize,
            int TaskIconSize,
            bool LazyLoadProgramsSubmenu,
            bool PlayVolumeFeedbackSound,
            bool UseClassicVolumePopup,
            int StartMenuSubmenuOpenDelayMs,
            int StartMenuRecentProgramsMaxCount,
            bool AutoHideTaskbar,
            bool StartOnWindowsStartup,
            TaskbarButtonStyle ButtonStyle,
            int BevelSize,
            string FontName,
            float FontSize,
            Color FontColor,
            Color BaseColor,
            Color LightColor,
            Color DarkColor,
            string ThemeProfileName);

        private sealed record ToolbarLayoutState(
            bool ShowQuickLaunch,
            bool ShowAddressInLeftRail,
            bool UseDedicatedAddressRow,
            bool ShowVolume,
            bool ShowSpotify);

        private sealed record ToolbarGripRelation(
            GripPanel Grip,
            Func<ToolbarLayoutState, bool> IsVisible);

        private sealed record ToolbarSeparatorRelation(
            Panel Separator,
            Func<ToolbarLayoutState, bool> IsVisible);

        private GripPanel CreateGripPanel(bool visible = false, bool enabled = true, Cursor? cursor = null, string? dragTag = null, bool enableToolbarReorder = false, bool forwardChildResizeMouse = false)
        {
            return new GripPanel(this, taskbarBaseColor, cursor ?? Cursors.SizeWE, visible, enabled, enableToolbarReorder, dragTag, forwardChildResizeMouse);
        }

        private Panel CreateSeparatorPanel()
        {
            var separatorPanel = new Panel
            {
                Dock = DockStyle.Left,
                Width = 4,
                Margin = Padding.Empty,
                Padding = Padding.Empty,
                BackColor = taskbarBaseColor
            };
            separatorPanel.Paint += ToolbarSeparatorPanel_Paint;
            return separatorPanel;
        }

        private sealed class Win9xColorTable : ProfessionalColorTable
        {
            private readonly Color baseColor;
            private readonly Color lightColor;
            private readonly Color darkColor;
            private readonly Color selectionColor;

            public Win9xColorTable(Color baseColor, Color lightColor, Color darkColor, Color selectionColor)
            {
                this.baseColor = baseColor;
                this.lightColor = lightColor;
                this.darkColor = darkColor;
                this.selectionColor = selectionColor;
            }

            public override Color ToolStripDropDownBackground => baseColor;
            public override Color ImageMarginGradientBegin => baseColor;
            public override Color ImageMarginGradientMiddle => baseColor;
            public override Color ImageMarginGradientEnd => baseColor;
            public override Color MenuItemSelected => selectionColor;
            public override Color MenuItemSelectedGradientBegin => selectionColor;
            public override Color MenuItemSelectedGradientEnd => selectionColor;
            public override Color MenuItemBorder => selectionColor;
            public override Color MenuBorder => darkColor;
            public override Color SeparatorDark => darkColor;
            public override Color SeparatorLight => lightColor;
        }

        private sealed class Win9xMenuRendererImpl : ToolStripProfessionalRenderer
        {
            private readonly Color menuBase;
            private readonly Color menuBlue;
            private readonly Color menuLight;
            private readonly Color menuDark;

            public Win9xMenuRendererImpl(Color baseColor, Color lightColor, Color darkColor, Color selectionColor)
                : base(new Win9xColorTable(baseColor, lightColor, darkColor, selectionColor))
            {
                menuBase = baseColor;
                menuLight = lightColor;
                menuDark = darkColor;
                menuBlue = selectionColor;
                RoundedEdges = false;
            }

            protected override void OnRenderToolStripBackground(ToolStripRenderEventArgs e)
            {
                using var brush = new SolidBrush(menuBase);
                e.Graphics.FillRectangle(brush, e.AffectedBounds);
            }

            protected override void OnRenderToolStripBorder(ToolStripRenderEventArgs e)
            {
                var rect = new Rectangle(0, 0, e.ToolStrip.Width - 1, e.ToolStrip.Height - 1);
                if (rect.Width <= 1 || rect.Height <= 1)
                {
                    return;
                }

                using var darkPen = new Pen(menuDark);
                using var lightPen = new Pen(menuLight);

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

                using var backBrush = new SolidBrush(menuBase);
                e.Graphics.FillRectangle(backBrush, rect);

                if (e.Item.Selected && e.Item.Enabled)
                {
                    var highlightRect = new Rectangle(1, 1, Math.Max(1, e.Item.Width - 2), Math.Max(1, e.Item.Height - 2));
                    using var highlightBrush = new SolidBrush(menuBlue);
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
                    TextRenderer.DrawText(e.Graphics, text, font, shadowRect, menuLight, flags);
                    TextRenderer.DrawText(e.Graphics, text, font, e.TextRectangle, menuDark, flags);
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

                using var darkPen = new Pen(menuDark);
                using var lightPen = new Pen(menuLight);
                e.Graphics.DrawLine(darkPen, left, y, right, y);
                e.Graphics.DrawLine(lightPen, left, y + 1, right, y + 1);
            }

            protected override void OnRenderItemCheck(ToolStripItemImageRenderEventArgs e)
            {
                if (e.Item is not ToolStripMenuItem menuItem || !menuItem.Checked)
                {
                    return;
                }

                var rect = e.ImageRectangle;
                if (rect.Width <= 0 || rect.Height <= 0)
                {
                    rect = new Rectangle(3, 3, 16, Math.Max(16, e.Item.Height - 6));
                }

                using var pen = new Pen(menuItem.Selected ? Color.White : Color.Black, 2f);
                var x1 = rect.Left + 3;
                var y1 = rect.Top + (rect.Height / 2);
                var x2 = rect.Left + (rect.Width / 2) - 1;
                var y2 = rect.Bottom - 4;
                var x3 = rect.Right - 3;
                var y3 = rect.Top + 4;
                e.Graphics.DrawLines(pen, new[] { new Point(x1, y1), new Point(x2, y2), new Point(x3, y3) });
            }

            // Note: Using native arrow rendering - no custom override needed
        }

        private int TaskbarButtonHeight => TaskbarRowHeight - TaskbarButtonVerticalPadding - TaskbarBottomPadding - 1;

        private int QuickLaunchButtonSize => TaskbarButtonHeight;

        private int ExpandedTaskbarHeight => taskbarRows * TaskbarRowHeight;

        private int ResizeGripHeight => Math.Max(ResizeGripMinHeight, taskbarBevelSize);

        private int TotalTaskbarHeight => ExpandedTaskbarHeight + ResizeGripHeight;

        protected override CreateParams CreateParams
        {
            get
            {
                var createParams = base.CreateParams;
                createParams.ExStyle |= WS_EX_TOOLWINDOW;
                createParams.ExStyle &= ~WS_EX_APPWINDOW;
                return createParams;
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (!allowAppExit && e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                return;
            }

            base.OnFormClosing(e);
        }

        private void RequestAppExit()
        {
            allowAppExit = true;
            Application.Exit();
        }

        public RetroTaskbarForm()
        {
            instance = this;
            shellManager = new ShellManager();
            interactionStateMachine = new TaskbarInteractionStateMachine(TimeSpan.FromMilliseconds(250));
            quickLaunchRefreshMonitor = new QuickLaunchRefreshMonitor(this);
            quickLaunchRefreshMonitor.RefreshRequested += (_, _) => LoadQuickLaunchButtons();

            var settings = TaskbarSettingsService.Load();
            startMenuIconSize = Math.Clamp(settings.StartMenuIconSize, 16, 32);
            taskIconSize = Math.Clamp(settings.TaskIconSize, 16, 32);
            lazyLoadProgramsSubmenu = settings.LazyLoadProgramsSubmenu;
            playVolumeFeedbackSound = settings.PlayVolumeFeedbackSound;
            useClassicVolumePopup = settings.UseClassicVolumePopup;
            startMenuSubmenuOpenDelayMs = Math.Clamp(settings.StartMenuSubmenuOpenDelayMs, 0, 1500);
            startMenuRecentProgramsMaxCount = Math.Clamp(settings.StartMenuRecentProgramsMaxCount, 1, 30);
            autoHideTaskbar = settings.AutoHideTaskbar;
            startOnWindowsStartup = settings.StartOnWindowsStartup;
            showQuickLaunchToolbar = settings.ShowQuickLaunchToolbar;
            showAddressToolbar = settings.ShowAddressToolbar;
            showSpotifyToolbar = settings.ShowSpotifyToolbar;
            showVolumeToolbar = settings.ShowVolumeToolbar;
            addressToolbarBeforeQuickLaunch = settings.AddressToolbarBeforeQuickLaunch;
            lastAddressToolbarText = settings.LastAddressText?.Trim() ?? string.Empty;
            addressToolbarHistory.Clear();
            addressToolbarHistory.AddRange(settings.AddressHistory);
            startMenuRecentPrograms.Clear();
            startMenuRecentPrograms.AddRange(settings.StartMenuRecentPrograms
                .Select(item => new ProgramSearchEntry(
                    string.IsNullOrWhiteSpace(item.DisplayName) ? item.LaunchPath : item.DisplayName,
                    item.LaunchPath.Trim(),
                    item.IsShellApp))
                .Take(startMenuRecentProgramsMaxCount)
                .ToList());
            taskbarFontName = settings.TaskbarFontName;
            taskbarFontSize = settings.TaskbarFontSize;
            taskbarFontColor = settings.TaskbarFontColor;
            taskbarLocked = settings.TaskbarLocked;
            taskbarRows = settings.TaskbarRows;
            taskbarBevelSize = settings.TaskbarBevelSize;
            taskbarBaseColor = settings.TaskbarBaseColor;
            taskbarLightColor = settings.TaskbarLightColor;
            taskbarDarkColor = settings.TaskbarDarkColor;
            selectedThemeProfileName = settings.ThemeProfileName;
            if (Enum.TryParse<TaskbarButtonStyle>(settings.TaskbarButtonStyle, ignoreCase: true, out var parsedStyle))
            {
                taskbarButtonStyle = parsedStyle;
            }

            ApplyThemeProfile(selectedThemeProfileName, preserveCustomizations: true);
            win9xMenuRenderer = CreateWin9xMenuRenderer();
            ApplyWindowsStartupSetting(startOnWindowsStartup);

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
            customVolumeIconImage = LoadVolumeTrayIconAsset(LoadAssetBySuffix("Volume.png", LoadMenuAsset("Audio Devices.png", SystemIcons.Information.ToBitmap())));
            spotifyPreviousIconImage = LoadPlayerControlIconAsset("previous.png", SystemIcons.Application.ToBitmap());
            spotifyPlayIconImage = LoadPlayerControlIconAsset("play.png", SystemIcons.Application.ToBitmap());
            spotifyNextIconImage = LoadPlayerControlIconAsset("next.png", SystemIcons.Application.ToBitmap());

            taskbarResizeHandlePanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = ResizeGripHeight,
                BackColor = taskbarBaseColor,
                Cursor = Cursors.SizeNS,
                Visible = true
            };
            taskbarResizeHandlePanel.Paint += ResizeGripPanel_Paint;

            contentPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = taskbarBaseColor,
                Padding = new Padding(0, TaskbarButtonVerticalPadding, 0, TaskbarBottomPadding)
            };


            startHostPanel = new Panel
            {
                Dock = DockStyle.Left,
                Width = 75,
                BackColor = taskbarBaseColor,
                Padding = new Padding(1, 0, TaskbarButtonGap, 0)
            };

            startToolbarSeparatorPanel = CreateSeparatorPanel();
            toolbarsSeparatorPanel = CreateSeparatorPanel();
            spotifyToolbarSeparatorPanel = CreateSeparatorPanel();
            taskButtonsSeparatorPanel = CreateSeparatorPanel();
            addressToolbarSeparatorPanel = CreateSeparatorPanel();

            addressRowSeparatorPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = AddressRowSeparatorHeight,
                Margin = Padding.Empty,
                Padding = Padding.Empty,
                BackColor = taskbarBaseColor,
                Visible = false
            };
            addressRowSeparatorPanel.Paint += AddressRowSeparatorPanel_Paint;


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
            lblClock.Paint += LblClock_Paint;

            languageIndicatorLabel = new Label
            {
                Width = 18,
                Height = 18,
                MinimumSize = new Size(18, 18),
                MaximumSize = new Size(18, 18),
                AutoSize = false,
                Margin = new Padding(1, 0, 1, 0),
                TextAlign = ContentAlignment.MiddleCenter,
                BorderStyle = BorderStyle.None,
                BackColor = InputIndicatorBackColor,
                ForeColor = Color.White,
                Font = InputIndicatorFont,
                Text = "En"
            };

            imeIndicatorLabel = new Label
            {
                Width = 18,
                Height = 18,
                MinimumSize = new Size(18, 18),
                MaximumSize = new Size(18, 18),
                AutoSize = false,
                Margin = new Padding(1, 0, 1, 0),
                TextAlign = ContentAlignment.MiddleCenter,
                BorderStyle = BorderStyle.None,
                BackColor = InputIndicatorBackColor,
                ForeColor = Color.White,
                Font = InputIndicatorFont,
                Text = "Al"
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

            quickLaunchGripPanel = CreateGripPanel(enableToolbarReorder: true, dragTag: "QuickLaunch", forwardChildResizeMouse: true);

            addressRowHostPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = TaskbarRowHeight,
                Margin = Padding.Empty,
                Padding = Padding.Empty,
                BackColor = taskbarBaseColor,
                BorderStyle = BorderStyle.None,
                Visible = false
            };

            addressToolbarPanel = new Panel
            {
                Dock = DockStyle.Left,
                Width = 260,
                Margin = Padding.Empty,
                Padding = new Padding(2, 1, 2, 2),
                BackColor = taskbarBaseColor,
                BorderStyle = BorderStyle.None,
                Visible = showAddressToolbar
            };


            volumeToolbarPanel = new Panel
            {
                Dock = DockStyle.Left,
                Width = 140,
                Margin = Padding.Empty,
                Padding = new Padding(6, 3, 6, 2),
                BackColor = taskbarBaseColor,
                BorderStyle = BorderStyle.None,
                Visible = true
            };

            volumeToolbarGripPanel = CreateGripPanel(forwardChildResizeMouse: true);
            addressToolbarGripPanel = CreateGripPanel(enableToolbarReorder: true, dragTag: "Address");
            addressRowGripPanel = CreateGripPanel();

            addressInputHostPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Margin = Padding.Empty,
                Padding = new Padding(2, 2, 2, 2),
                BackColor = Color.White,
                BorderStyle = BorderStyle.None
            };
            addressInputHostPanel.Paint += AddressInputHostPanel_Paint;

            addressToolbarComboBox = new ComboBox
            {
                Dock = DockStyle.None,
                DropDownStyle = ComboBoxStyle.DropDown,
                FlatStyle = FlatStyle.Flat,
                Font = normalFont,
                Margin = Padding.Empty,
                IntegralHeight = false,
                MaxDropDownItems = 12,
                AutoCompleteMode = AutoCompleteMode.SuggestAppend,
                AutoCompleteSource = AutoCompleteSource.ListItems,
                BackColor = Color.White,
                ForeColor = taskbarFontColor,
                Text = string.Empty
            };
            addressToolbarComboBox.HandleCreated += (_, _) => ApplyClassicAddressComboStyle();
            addressToolbarComboBox.KeyDown += AddressToolbarComboBox_KeyDown;
            addressToolbarComboBox.Anchor = AnchorStyles.Left | AnchorStyles.Right;
            addressToolbarComboBox.Height = Math.Max(18, TextRenderer.MeasureText("My Computer", normalFont).Height + 8);

            volumeToolbarTrackBar = new TrackBar
            {
                Dock = DockStyle.Fill,
                Minimum = 0,
                Maximum = 100,
                TickStyle = TickStyle.None,
                SmallChange = 2,
                LargeChange = 10,
                Margin = Padding.Empty
            };
            volumeToolbarTrackBar.ValueChanged += VolumeToolbarTrackBar_ValueChanged;
            volumeToolbarTrackBar.MouseDown += VolumeToolbarTrackBar_MouseDown;
            volumeToolbarTrackBar.MouseUp += VolumeToolbarTrackBar_MouseUp;

            addressInputHostPanel.Resize += (_, _) => CenterAddressTextBoxVertically();
            addressInputHostPanel.Controls.Add(addressToolbarComboBox);
            addressInputHostPanel.Height = GetAddressInputHostHeight();
            RefreshAddressToolbarHistoryItems();
            if (!string.IsNullOrWhiteSpace(lastAddressToolbarText))
            {
                addressToolbarComboBox.Text = lastAddressToolbarText;
                addressToolbarComboBox.SelectionStart = addressToolbarComboBox.Text.Length;
            }
            CenterAddressTextBoxVertically();
            volumeToolbarPanel.Controls.Add(volumeToolbarTrackBar);
            SyncVolumeToolbarFromSystem();

            addressToolbarPanel.Controls.Add(addressInputHostPanel);
            addressRowHostPanel.Controls.Add(addressRowSeparatorPanel);


            addressRowHostPanel.Controls.Add(volumeToolbarGripPanel);
            addressRowHostPanel.Controls.Add(volumeToolbarPanel);
            
            addressRowHostPanel.Controls.Add(addressToolbarSeparatorPanel);
            addressRowHostPanel.Controls.Add(addressToolbarPanel);

            addressRowHostPanel.Controls.Add(addressRowGripPanel);
            
            spotifyToolbarPanel = new Panel
            {
                Dock = DockStyle.Left,
                Width = (QuickLaunchButtonSize * 3) + (TaskbarButtonGap * 2) + 4,
                Margin = Padding.Empty,
                Padding = new Padding(2, 0, 2, 3),
                BackColor = taskbarBaseColor,
                BorderStyle = BorderStyle.None,
                Visible = showSpotifyToolbar
            };

            spotifyControlsPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = QuickLaunchButtonSize,
                Margin = Padding.Empty,
                Padding = Padding.Empty,
                BackColor = taskbarBaseColor
            };

            spotifyToolbarGripPanel = CreateGripPanel(forwardChildResizeMouse: true);
            quickLaunchToolbar = new ToolbarComponents(quickLaunchPanel, () => quickLaunchGripPanel);
            addressToolbar = new ToolbarComponents(addressToolbarPanel, () => addressToolbarGripPanel);
            volumeToolbar = new ToolbarComponents(volumeToolbarPanel, () => volumeToolbarGripPanel);
            spotifyToolbar = new ToolbarComponents(spotifyToolbarPanel, () => spotifyToolbarGripPanel);
            toolbarsById = new Dictionary<ToolbarId, ToolbarComponents>
            {
                [ToolbarId.QuickLaunch] = quickLaunchToolbar,
                [ToolbarId.Address] = addressToolbar,
                [ToolbarId.Volume] = volumeToolbar,
                [ToolbarId.Spotify] = spotifyToolbar
            };

            btnSpotifyPrevious = new Button
            {
                Width = QuickLaunchButtonSize,
                Height = QuickLaunchButtonSize,
                Dock = DockStyle.Left,
                Image = ResizeIconImage(spotifyPreviousIconImage, GetSpotifyControlIconSize()),
                ImageAlign = ContentAlignment.MiddleCenter,
                Text = string.Empty,
                Font = normalFont,
                Margin = new Padding(0, 0, 2, 0),
                Padding = Padding.Empty,
                TabStop = false
            };
            btnSpotifyPrevious.Click += (_, _) => ExecuteSpotifyCommand(SpotifyController.SpotifyCommand.PreviousTrack);
            ApplyQuickLaunchButtonStyle(btnSpotifyPrevious);
            AttachNoFocusCue(btnSpotifyPrevious);

            btnSpotifyPlayPause = new Button
            {
                Width = QuickLaunchButtonSize,
                Height = QuickLaunchButtonSize,
                Dock = DockStyle.Left,
                Image = ResizeIconImage(spotifyPlayIconImage, GetSpotifyControlIconSize()),
                ImageAlign = ContentAlignment.MiddleCenter,
                Text = string.Empty,
                Font = normalFont,
                Margin = new Padding(0, 0, 2, 0),
                Padding = Padding.Empty,
                TabStop = false
            };
            btnSpotifyPlayPause.Click += (_, _) => ExecuteSpotifyCommand(SpotifyController.SpotifyCommand.PlayPause);
            ApplyQuickLaunchButtonStyle(btnSpotifyPlayPause);
            AttachNoFocusCue(btnSpotifyPlayPause);

            btnSpotifyNext = new Button
            {
                Width = QuickLaunchButtonSize,
                Height = QuickLaunchButtonSize,
                Dock = DockStyle.Left,
                Image = ResizeIconImage(spotifyNextIconImage, GetSpotifyControlIconSize()),
                ImageAlign = ContentAlignment.MiddleCenter,
                Text = string.Empty,
                Font = normalFont,
                Margin = Padding.Empty,
                Padding = Padding.Empty,
                TabStop = false
            };
            btnSpotifyNext.Click += (_, _) => ExecuteSpotifyCommand(SpotifyController.SpotifyCommand.NextTrack);
            ApplyQuickLaunchButtonStyle(btnSpotifyNext);
            AttachNoFocusCue(btnSpotifyNext);

            spotifyNowPlayingLabel = new Label
            {
                Dock = DockStyle.Fill,
                Height = 18,
                TextAlign = ContentAlignment.MiddleLeft,
                AutoEllipsis = true,
                Margin = Padding.Empty,
                Padding = new Padding(3, 0, 0, 0),
                BackColor = taskbarBaseColor,
                ForeColor = taskbarFontColor,
                Font = normalFont,
                Visible = false
            };

            spotifyControlsPanel.Controls.Add(btnSpotifyNext);
            spotifyControlsPanel.Controls.Add(btnSpotifyPlayPause);
            spotifyControlsPanel.Controls.Add(btnSpotifyPrevious);
            spotifyToolbarPanel.Controls.Add(spotifyNowPlayingLabel);
            spotifyToolbarPanel.Controls.Add(spotifyControlsPanel);
            isSpotifyRunning = spotifyController.IsSpotifyRunning();
            btnSpotifyPrevious.Enabled = true;
            btnSpotifyPlayPause.Enabled = true;
            btnSpotifyNext.Enabled = true;

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
                FlowDirection = FlowDirection.LeftToRight,
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

            runningProgramsGripPanel = CreateGripPanel(
                visible: true,
                enabled: false,
                cursor: Cursors.Default);

            toolbarGripRelations.Add(new ToolbarGripRelation(quickLaunchGripPanel, state => state.ShowQuickLaunch));
            toolbarGripRelations.Add(new ToolbarGripRelation(addressToolbarGripPanel, state => state.ShowAddressInLeftRail));
            toolbarGripRelations.Add(new ToolbarGripRelation(addressRowGripPanel, state => state.UseDedicatedAddressRow));
            toolbarGripRelations.Add(new ToolbarGripRelation(volumeToolbarGripPanel, state => state.ShowVolume));
            toolbarGripRelations.Add(new ToolbarGripRelation(spotifyToolbarGripPanel, state => state.ShowSpotify));
            toolbarGripRelations.Add(new ToolbarGripRelation(runningProgramsGripPanel, _ => true));

            toolbarSeparatorRelations.Add(new ToolbarSeparatorRelation(
                toolbarsSeparatorPanel,
                state => state.ShowQuickLaunch && state.ShowAddressInLeftRail));
            toolbarSeparatorRelations.Add(new ToolbarSeparatorRelation(
                spotifyToolbarSeparatorPanel,
                state => state.ShowSpotify && (state.ShowQuickLaunch || state.ShowAddressInLeftRail || state.ShowVolume)));
            toolbarSeparatorRelations.Add(new ToolbarSeparatorRelation(
                startToolbarSeparatorPanel,
                state => state.ShowQuickLaunch || state.ShowAddressInLeftRail || state.ShowVolume || state.ShowSpotify));

            taskButtonsHostPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Margin = Padding.Empty,
                Padding = Padding.Empty,
                BackColor = taskbarBaseColor,
                BorderStyle = BorderStyle.None
            };
            taskButtonsHostPanel.Controls.Add(taskButtonsPanel);
            taskButtonsHostPanel.Controls.Add(runningProgramsGripPanel);
            taskButtonsHostPanel.Controls.Add(addressRowHostPanel);

            startMenu = BuildStartMenu();
            startMenuController = new StartMenuController(
                interactionStateMachine,
                () => startMenu.Visible,
                ToggleStartMenu,
                CloseStartMenuIfVisible);
            taskWindowMenu = BuildTaskWindowMenu();
            taskbarContextMenu = BuildTaskbarContextMenu();
            clockContextMenu = BuildClockContextMenu();
            languageIndicatorContextMenu = BuildLanguageIndicatorContextMenu();
            imeIndicatorContextMenu = BuildImeIndicatorContextMenu();
            quickLaunchItemContextMenu = BuildQuickLaunchItemContextMenu();
            taskToolTip = new ToolTip();
            trayToolTip = new ToolTip
            {
                ShowAlways = true,
                InitialDelay = 500,
                ReshowDelay = 100,
                AutoPopDelay = 5000,
                UseAnimation = false,
                UseFading = false
            };
            quickLaunchToolTip = new ToolTip();
            keyboardProc = KeyboardHookCallback;
            mouseProc = MouseHookCallback;
            volumePopup = new VolumePopupForm();
            volumePopup.PlayFeedbackSoundOnMouseUp = playVolumeFeedbackSound;
            timeDetailsPopup = new TimeDetailsForm();
            LoadQuickLaunchButtons();

            startHostPanel.Controls.Add(btnStart);
            notificationAreaPanel.Controls.Add(lblClock);
            notificationAreaPanel.Controls.Add(trayIconsPanel);
            contentPanel.Controls.Add(taskButtonsHostPanel);
            contentPanel.Controls.Add(notificationAreaPanel);
            contentPanel.Controls.Add(startToolbarSeparatorPanel);
            contentPanel.Controls.Add(startHostPanel);

            Controls.Add(contentPanel);
            Controls.Add(taskbarResizeHandlePanel);
            trayIconsPanel.Controls.Add(customVolumeIcon);
            trayIconsPanel.Controls.Add(languageIndicatorLabel);
            trayToolTip.SetToolTip(customVolumeIcon, "Volume");
            trayToolTip.SetToolTip(languageIndicatorLabel, "Keyboard language");
            quickLaunchToolTip.SetToolTip(btnSpotifyPrevious, "Spotify Previous");
            quickLaunchToolTip.SetToolTip(btnSpotifyPlayPause, "Spotify Play/Pause");
            quickLaunchToolTip.SetToolTip(btnSpotifyNext, "Spotify Next");
            UpdateInputMethodIndicators();
            ApplyToolbarLayout(refreshBounds: false);
            UpdateResizeUiState();

            MouseUp += TaskbarBackground_MouseUp;
            MouseDown += TaskbarResize_MouseDown;
            MouseMove += TaskbarResize_MouseMove;
            MouseUp += TaskbarResize_MouseUp;
            taskbarResizeHandlePanel.MouseDown += TaskbarResize_MouseDown;
            taskbarResizeHandlePanel.MouseMove += TaskbarResize_MouseMove;
            taskbarResizeHandlePanel.MouseUp += TaskbarResize_MouseUp;
            taskButtonsPanel.MouseUp += TaskButtonsPanel_MouseUp;
            taskButtonsPanel.MouseDown += ChildTaskbarResize_MouseDown;
            taskButtonsPanel.MouseMove += ChildTaskbarResize_MouseMove;
            taskButtonsPanel.MouseUp += ChildTaskbarResize_MouseUp;
            lblClock.MouseUp += LblClock_MouseUp;
            lblClock.MouseClick += LblClock_MouseClick;
            lblClock.MouseDown += ChildTaskbarResize_MouseDown;
            lblClock.MouseMove += ChildTaskbarResize_MouseMove;
            lblClock.MouseUp += ChildTaskbarResize_MouseUp;
            languageIndicatorLabel.MouseUp += InputIndicatorLabel_MouseUp;
            imeIndicatorLabel.MouseUp += InputIndicatorLabel_MouseUp;
            quickLaunchPanel.MouseDown += ChildTaskbarResize_MouseDown;
            quickLaunchPanel.MouseMove += ChildTaskbarResize_MouseMove;
            quickLaunchPanel.MouseUp += ChildTaskbarResize_MouseUp;
            spotifyToolbarPanel.MouseDown += ChildTaskbarResize_MouseDown;
            spotifyToolbarPanel.MouseMove += ChildTaskbarResize_MouseMove;
            spotifyToolbarPanel.MouseUp += ChildTaskbarResize_MouseUp;
            volumeToolbarPanel.MouseDown += ChildTaskbarResize_MouseDown;
            volumeToolbarPanel.MouseMove += ChildTaskbarResize_MouseMove;
            volumeToolbarPanel.MouseUp += ChildTaskbarResize_MouseUp;
            btnSpotifyPrevious.MouseDown += ChildTaskbarResize_MouseDown;
            btnSpotifyPrevious.MouseMove += ChildTaskbarResize_MouseMove;
            btnSpotifyPrevious.MouseUp += ChildTaskbarResize_MouseUp;
            btnSpotifyPlayPause.MouseDown += ChildTaskbarResize_MouseDown;
            btnSpotifyPlayPause.MouseMove += ChildTaskbarResize_MouseMove;
            btnSpotifyPlayPause.MouseUp += ChildTaskbarResize_MouseUp;
            btnSpotifyNext.MouseDown += ChildTaskbarResize_MouseDown;
            btnSpotifyNext.MouseMove += ChildTaskbarResize_MouseMove;
            btnSpotifyNext.MouseUp += ChildTaskbarResize_MouseUp;
            spotifyNowPlayingLabel.MouseDown += ChildTaskbarResize_MouseDown;
            spotifyNowPlayingLabel.MouseMove += ChildTaskbarResize_MouseMove;
            spotifyNowPlayingLabel.MouseUp += ChildTaskbarResize_MouseUp;
            trayIconsPanel.MouseDown += ChildTaskbarResize_MouseDown;
            trayIconsPanel.MouseMove += ChildTaskbarResize_MouseMove;
            trayIconsPanel.MouseUp += ChildTaskbarResize_MouseUp;
            quickLaunchPanel.DragEnter += QuickLaunchPanel_DragEnter;
            quickLaunchPanel.DragOver += QuickLaunchPanel_DragOver;
            quickLaunchPanel.DragLeave += QuickLaunchPanel_DragLeave;
            quickLaunchPanel.DragDrop += QuickLaunchPanel_DragDrop;
            quickLaunchPanel.Paint += QuickLaunchPanel_Paint;

            refreshTimer = new System.Windows.Forms.Timer { Interval = RefreshIntervalMs };
            refreshTimer.Tick += (_, _) => RefreshTaskbar();
            trayIconStabilityTimer = new System.Windows.Forms.Timer { Interval = TrayIconNewIconStabilityMs };
            trayIconStabilityTimer.Tick += (_, _) =>
            {
                trayIconStabilityTimer.Stop();
                SyncTrayIcons();
            };
            trayResizeShrinkTimer = new System.Windows.Forms.Timer { Interval = 900 };
            trayResizeShrinkTimer.Tick += (_, _) =>
            {
                trayResizeShrinkTimer.Stop();
                ResizeTrayPanel(allowShrink: true);
                RefreshTrayHostPlacement();
            };

            Load += RetroTaskbarForm_Load;
            FormClosed += RetroTaskbarForm_FormClosed;
            Paint += RetroTaskbarForm_Paint;
            SizeChanged += (_, _) => RefreshTaskButtonAndTrayLayout();
            LocationChanged += (_, _) => RefreshTrayHostPlacement();
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WinApi.WM_MOUSEACTIVATE)
            {
                m.Result = (IntPtr)WinApi.MA_NOACTIVATE;
                return;
            }

            if (m.Msg == (int)WinApi.AppBarMsg.WindowPosChanged)
            {
                // Re-register AppBar when window position changes
                RegisterAppBar();
                m.Result = IntPtr.Zero;
                return;
            }

            // Handle AppBar callback message (fullscreen notifications, etc.)
            if (m.Msg == (int)WinApi.WM_USER + 0x7FFF)
            {
                var notifyCode = m.WParam.ToInt32();
                if (notifyCode == WinApi.ABN_FULLSCREENAPPNOTIFY)
                {
                    HandleFullscreenNotify(m.LParam);
                }
                else if (notifyCode == WinApi.ABN_POSCHANGED)
                {
                    // AppBar position changed - may need to re-register
                    RegisterAppBar();
                }
                m.Result = IntPtr.Zero;
                return;
            }

            base.WndProc(ref m);
        }

        private void RegisterAppBar()
        {
            if (isAppBarRegistered)
            {
                UpdateAppBarPosition();
                return;
            }

            var appBarData = new WinApi.APPBARDATA
            {
                cbSize = (uint)Marshal.SizeOf<WinApi.APPBARDATA>(),
                hWnd = Handle,
                uCallbackMessage = (uint)WinApi.WM_USER + 0x7FFF, // Custom message
                uEdge = (uint)WinApi.AppBarEdge.Bottom,
                rc = new WinApi.RECT(0, 0, 0, 0),
                lParam = IntPtr.Zero
            };

            var result = WinApi.SHAppBarMessage((uint)WinApi.AppBarMsg.New, ref appBarData);
            appBarHandle = result;

            if (result != IntPtr.Zero)
            {
                isAppBarRegistered = true;
                UpdateAppBarPosition();
            }
        }

        private void UnregisterAppBar()
        {
            if (isAppBarRegistered && appBarHandle != IntPtr.Zero)
            {
                var appBarData = new WinApi.APPBARDATA
                {
                    cbSize = (uint)Marshal.SizeOf<WinApi.APPBARDATA>(),
                    hWnd = Handle,
                    uCallbackMessage = (uint)WinApi.WM_USER + 0x7FFF,
                    uEdge = (uint)WinApi.AppBarEdge.Bottom,
                    rc = new WinApi.RECT(0, 0, 0, 0),
                    lParam = IntPtr.Zero
                };

                WinApi.SHAppBarMessage((uint)WinApi.AppBarMsg.Remove, ref appBarData);
                appBarHandle = IntPtr.Zero;
                isAppBarRegistered = false;
            }
        }

        private void HandleFullscreenNotify(IntPtr lParam)
        {
            // ABN_FULLSCREENAPPNOTIFY: lParam = TRUE means entering fullscreen, FALSE means exiting
            var enteringFullscreen = lParam.ToInt32() != 0;
            if (enteringFullscreen && IsAltKeyDown())
            {
                return;
            }

            suppressFullscreenAutoDetect = true;
            SetFullscreenState(enteringFullscreen);
        }

        private void SetFullscreenState(bool isFullscreen)
        {
            if (isFullscreen && !isFullscreenActive)
            {
                isFullscreenActive = true;
                if (Visible)
                {
                    Hide();
                }
            }
            else if (!isFullscreen && isFullscreenActive)
            {
                isFullscreenActive = false;
                if (!Visible)
                {
                    Show();
                }
            }
        }

        private void DetectFullscreenWindow()
        {
            if (IsAltKeyDown())
            {
                return;
            }

            if (suppressFullscreenAutoDetect)
            {
                suppressFullscreenAutoDetect = false;
                return;
            }

            var foregroundHandle = WinApi.GetForegroundWindow();
            if (foregroundHandle == IntPtr.Zero || foregroundHandle == Handle || WinApi.IsIconic(foregroundHandle))
            {
                SetFullscreenState(false);
                return;
            }

            if (IsDesktopOrShellWindow(foregroundHandle))
            {
                SetFullscreenState(false);
                return;
            }

            if (!WinApi.GetWindowRect(foregroundHandle, out var foregroundRect))
            {
                SetFullscreenState(false);
                return;
            }

            var screenBounds = Screen.FromHandle(foregroundHandle).Bounds;
            const int tolerance = 2;
            var isFullscreen =
                foregroundRect.Left <= screenBounds.Left + tolerance &&
                foregroundRect.Top <= screenBounds.Top + tolerance &&
                foregroundRect.Right >= screenBounds.Right - tolerance &&
                foregroundRect.Bottom >= screenBounds.Bottom - tolerance;

            SetFullscreenState(isFullscreen);
        }

        private static bool IsDesktopOrShellWindow(IntPtr windowHandle)
        {
            if (windowHandle == WinApi.GetShellWindow())
            {
                return true;
            }

            var classNameBuffer = new StringBuilder(256);
            if (WinApi.GetClassName(windowHandle, classNameBuffer, classNameBuffer.Capacity) == 0)
            {
                return false;
            }

            var className = classNameBuffer.ToString();
            return className.Equals("Progman", StringComparison.OrdinalIgnoreCase) ||
                   className.Equals("WorkerW", StringComparison.OrdinalIgnoreCase) ||
                   className.Equals("Shell_TrayWnd", StringComparison.OrdinalIgnoreCase) ||
                   className.Equals("Shell_SecondaryTrayWnd", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsAltKeyDown() => (WinApi.GetAsyncKeyState(WinApi.VK_MENU) & 0x8000) != 0;

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
            SystemEvents.DisplaySettingsChanged += SystemEvents_DisplaySettingsChanged;
            InstallKeyboardHook();
            InstallMouseHook();
            InitializeTrayService();
            quickLaunchRefreshMonitor.Start();
            RegisterAppBar(); // Register as AppBar to reserve desktop space
            SetTaskbarBounds();
            RefreshTaskbar();
            refreshTimer.Start();
        }

        private void RetroTaskbarForm_FormClosed(object? sender, FormClosedEventArgs e)
        {
            instance = null;
            refreshTimer.Stop();
            refreshTimer.Dispose();
            trayIconStabilityTimer.Stop();
            trayIconStabilityTimer.Dispose();
            trayResizeShrinkTimer.Stop();
            trayResizeShrinkTimer.Dispose();
            SaveTaskbarSettings();
            SystemEvents.DisplaySettingsChanged -= SystemEvents_DisplaySettingsChanged;

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
            quickLaunchPanel.DragOver -= QuickLaunchPanel_DragOver;
            quickLaunchPanel.DragLeave -= QuickLaunchPanel_DragLeave;
            quickLaunchPanel.DragDrop -= QuickLaunchPanel_DragDrop;
            quickLaunchPanel.Paint -= QuickLaunchPanel_Paint;
            menuDropDownOpeningHandlers.Clear();

            btnStart.Image?.Dispose();
            startMenu.Dispose();
            taskWindowMenu.Dispose();
            taskbarContextMenu.Dispose();
            clockContextMenu.Dispose();
            languageIndicatorContextMenu.Dispose();
            imeIndicatorContextMenu.Dispose();
            quickLaunchItemContextMenu.Dispose();
            diagnosticsPanel?.Close();
            diagnosticsPanel?.Dispose();
            diagnosticsPanel = null;
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
            languageIndicatorLabel.MouseUp -= InputIndicatorLabel_MouseUp;
            imeIndicatorLabel.MouseUp -= InputIndicatorLabel_MouseUp;
            btnSpotifyPrevious.Image?.Dispose();
            btnSpotifyPlayPause.Image?.Dispose();
            btnSpotifyNext.Image?.Dispose();
            customVolumeIcon.Image?.Dispose();
            customVolumeIcon.Dispose();
            customVolumeIconImage.Dispose();
            volumeToolbarTrackBar.ValueChanged -= VolumeToolbarTrackBar_ValueChanged;
            volumeToolbarTrackBar.MouseDown -= VolumeToolbarTrackBar_MouseDown;
            volumeToolbarTrackBar.MouseUp -= VolumeToolbarTrackBar_MouseUp;
            spotifyPreviousIconImage.Dispose();
            spotifyPlayIconImage.Dispose();
            spotifyNextIconImage.Dispose();
            volumePopup.Dispose();
            timeDetailsPopup?.Dispose();
            timeDetailsPopup = null;
            UninstallKeyboardHook();
            UninstallMouseHook();
            UnregisterAppBar(); // Unregister AppBar before cleanup
            quickLaunchRefreshMonitor.Dispose();
            TeardownTrayService();
            lock (shellIconCacheLock)
            {
                foreach (var cachedIcon in shellIconCache.Values)
                {
                    cachedIcon.Dispose();
                }
                shellIconCache.Clear();
            }

            TrySetExplorerTaskbarHidden(false);
            shellManager.AppBarManager.SignalGracefulShutdown();
            shellManager.Dispose();
        }

        private void SystemEvents_DisplaySettingsChanged(object? sender, EventArgs e)
        {
            if (IsDisposed || Disposing)
            {
                return;
            }

            if (InvokeRequired)
            {
                BeginInvoke(new Action(RefreshTaskButtonAndTrayLayout));
                return;
            }

            RefreshTaskButtonAndTrayLayout();
        }

        private void RetroTaskbarForm_Paint(object? sender, PaintEventArgs e)
        {
            // Top bevel is rendered by taskbarResizeHandlePanel.
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
                ReadOnly = false
            };
            startMenuSearchTextBox.TextChanged += StartMenuSearchTextBox_TextChanged;
            startMenuSearchTextBox.KeyDown += StartMenuSearchTextBox_KeyDown;
            menu.Items.Add(startMenuSearchTextBox);
            menu.Items.Add(new ToolStripSeparator());
            RefreshStartMenuRecentProgramsSection(menu);

            menu.Items.Add(CreateMenuItem(
                "File Explorer",
                explorerMenuIcon,
                (_, _) => BringExplorerToFront(),
                new ProgramSearchEntry("File Explorer", "explorer.exe", IsShellApp: false)));

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

            var windowsAppsItem = CreateMenuItem("Windows Apps", LoadMenuAsset("Application Window.png", SystemIcons.Application.ToBitmap()), null);
            var windowsAppsLoadingItem = CreateMenuItem("(Loading...)", null, null);
            windowsAppsLoadingItem.Enabled = false;
            windowsAppsItem.DropDownItems.Add(windowsAppsLoadingItem);
            windowsAppsItem.DropDownOpening += (_, _) => PopulateWindowsAppsMenuAsync(windowsAppsItem);
            menu.Items.Add(windowsAppsItem);

            var documentsItem = CreateMenuItem("Documents", documentsMenuIcon, null);
            var documentsLoadingItem = CreateMenuItem("(Loading...)", null, null);
            documentsLoadingItem.Enabled = false;
            documentsItem.DropDownItems.Add(documentsLoadingItem);
            documentsItem.DropDownOpening += (_, _) => PopulateSubmenuWithOptionalDelay(documentsItem, () => PopulateDocumentsMenu(documentsItem.DropDownItems));
            menu.Items.Add(documentsItem);

            var systemToolsItem = CreateMenuItem("System Tools", LoadMenuAsset("Computer Management.png", SystemIcons.Application.ToBitmap()), null);
            PopulateSystemToolsMenu(systemToolsItem.DropDownItems);
            menu.Items.Add(systemToolsItem);

            menu.Items.Add(CreateMenuItem(
                "Run...",
                runMenuIcon,
                (_, _) => LaunchProcess("rundll32.exe", "shell32.dll,#61"),
                new ProgramSearchEntry("Run...", EncodeCommandLaunchPath("rundll32.exe", "shell32.dll,#61"), IsShellApp: false)));
            menu.Items.Add(CreateMenuItem("Invalidate Start Menu Cache", null, (_, _) => InvalidateStartMenuCache()));
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add(CreateMenuItem("Shut Down...", SystemIcons.Error.ToBitmap(), (_, _) => ShowShutdownPrompt()));

            menu.Closed += (_, _) =>
            {
                startMenuController.OnStartMenuClosed();
                ResetStartMenuSearch();
                ResetStartMenuSubmenuSessionCache();
                btnStart.Invalidate();
            };

            return menu;
        }

        private ToolStripMenuItem CreateMenuItem(string text, Image? image, EventHandler? onClick, ProgramSearchEntry? recentEntry = null)
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
                    if (recentEntry != null)
                    {
                        RecordStartMenuRecentProgram(recentEntry);
                    }
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

        private void AttachProgramSearchContextMenu(ToolStripMenuItem item, ProgramSearchEntry entry, bool trackRecentUsage = false)
        {
            item.MouseUp += (_, e) =>
            {
                if (e.Button != MouseButtons.Right)
                {
                    return;
                }

                using var menu = new ContextMenuStrip();
                ApplyWin9xContextMenuStyle(menu);

                var openItem = new ToolStripMenuItem("Open");
                openItem.Click += (_, _) => LaunchProgramSearchEntry(entry, trackRecentUsage);
                menu.Items.Add(openItem);

                if (!entry.IsShellApp && File.Exists(entry.LaunchPath))
                {
                    var openFileLocationItem = new ToolStripMenuItem("Open file location");
                    openFileLocationItem.Click += (_, _) => LaunchProcess("explorer.exe", $"/select,\"{entry.LaunchPath}\"");
                    menu.Items.Add(openFileLocationItem);
                }
                else if (entry.IsShellApp)
                {
                    var openAppsFolderItem = new ToolStripMenuItem("Open Apps folder");
                    openAppsFolderItem.Click += (_, _) => LaunchProcess("explorer.exe", "shell:AppsFolder");
                    menu.Items.Add(openAppsFolderItem);
                }

                menu.Show(Cursor.Position);
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
                var item = CreateMenuItem(
                    name,
                    GetSmallShellIcon(file, isDirectory: false),
                    (_, _) => LaunchProcess(file),
                    new ProgramSearchEntry(name, file, IsShellApp: false));
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
            lock (programSearchEntriesCacheLock)
            {
                programFolderCache.Clear();
                programSearchEntriesCache.Clear();
            }
        }

        private void InvalidateStartMenuCache()
        {
            InvalidateProgramMenuCache();
            ResetStartMenuSubmenuSessionCache();
            populatedSubmenusThisSession.RemoveWhere(item => string.Equals(item.Text, "Windows Apps", StringComparison.OrdinalIgnoreCase));
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

            // Ensure the dropdown uses our custom renderer
            SetupMenuItemRenderer(item);

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
            items.Add(CreateMenuItem(
                "My Documents",
                documentsMenuIcon,
                (_, _) => OpenFolder(myDocuments),
                new ProgramSearchEntry("My Documents", myDocuments, IsShellApp: false)));
            AddKnownFolderMenuItem(items, "Desktop", Environment.SpecialFolder.DesktopDirectory, LoadMenuAsset("Desktop.png", SystemIcons.Application.ToBitmap()));
            AddKnownFolderMenuItem(items, "Downloads", "Downloads", LoadMenuAsset("Folder Closed.png", SystemIcons.Application.ToBitmap()));
            AddKnownFolderMenuItem(items, "User Profile", Environment.SpecialFolder.UserProfile, LoadMenuAsset("User Accounts.png", SystemIcons.Application.ToBitmap()));

            var recent = Environment.GetFolderPath(Environment.SpecialFolder.Recent);
            if (Directory.Exists(recent))
            {
                var recents = Directory.GetFiles(recent, "*.lnk").OrderByDescending(File.GetLastWriteTime).Take(8).ToList();
                if (recents.Count > 0)
                {
                    items.Add(new ToolStripSeparator());
                    foreach (var path in recents)
                    {
                        var displayName = Path.GetFileNameWithoutExtension(path);
                        items.Add(CreateMenuItem(
                            displayName,
                            null,
                            (_, _) => LaunchProcess(path),
                            new ProgramSearchEntry(displayName, path, IsShellApp: false)));
                    }
                }
            }
        }

        private void PopulateWindowsAppsMenuAsync(ToolStripMenuItem windowsAppsItem)
        {
            if (populatedSubmenusThisSession.Contains(windowsAppsItem))
            {
                return;
            }

            populatedSubmenusThisSession.Add(windowsAppsItem);
            windowsAppsItem.DropDownItems.Clear();
            var loadingItem = CreateMenuItem("(Loading...)", null, null);
            loadingItem.Enabled = false;
            windowsAppsItem.DropDownItems.Add(loadingItem);

            _ = Task.Run(() =>
            {
                var apps = GetProgramSearchEntries()
                    .Where(entry => entry.IsShellApp)
                    .OrderBy(entry => entry.DisplayName, StringComparer.CurrentCultureIgnoreCase)
                    .ToList();
                var appItems = apps
                    .Select(app => new
                    {
                        App = app,
                        Icon = GetProgramSearchEntryIcon(app)
                    })
                    .ToList();

                if (IsDisposed || Disposing)
                {
                    foreach (var prepared in appItems)
                    {
                        prepared.Icon?.Dispose();
                    }
                    return;
                }

                BeginInvoke(new Action(() =>
                {
                    if (IsDisposed || Disposing || !windowsAppsItem.Visible)
                    {
                        foreach (var prepared in appItems)
                        {
                            prepared.Icon?.Dispose();
                        }
                        return;
                    }

                    var wasDropDownVisible = windowsAppsItem.DropDown.Visible;
                    if (wasDropDownVisible)
                    {
                        windowsAppsItem.HideDropDown();
                    }

                    windowsAppsItem.DropDown.SuspendLayout();
                    windowsAppsItem.DropDownItems.Clear();

                    if (appItems.Count == 0)
                    {
                        var emptyItem = CreateMenuItem("(No Windows Apps found)", null, null);
                        emptyItem.Enabled = false;
                        windowsAppsItem.DropDownItems.Add(emptyItem);
                    }
                    else
                    {
                        foreach (var prepared in appItems)
                        {
                            var item = CreateMenuItem(prepared.App.DisplayName, prepared.Icon, (_, _) => LaunchProgramSearchEntry(prepared.App), prepared.App);
                            AttachProgramSearchContextMenu(item, prepared.App, trackRecentUsage: true);
                            windowsAppsItem.DropDownItems.Add(item);
                            prepared.Icon?.Dispose();
                        }

                        ApplyTwoColumnLayoutIfNeeded(windowsAppsItem.DropDownItems);
                    }
                    windowsAppsItem.DropDown.ResumeLayout(performLayout: true);

                    if (wasDropDownVisible || !windowsAppsItem.DropDown.Visible)
                    {
                        windowsAppsItem.ShowDropDown();
                    }
                }));
            });
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
            var arrowSpace = items.Cast<ToolStripItem>().Any(item => item.Available && item is ToolStripMenuItem menuItem && menuItem.HasDropDownItems) ? 20 : 0;
            var itemWidth = Math.Clamp(maxMeasuredWidth + imageSpace + arrowSpace, 170, 380);
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

        private static string EncodeCommandLaunchPath(string fileName, string? arguments)
        {
            return $"{StartMenuCommandTokenPrefix}{fileName}|{arguments ?? string.Empty}";
        }

        private static bool TryDecodeCommandLaunchPath(string launchPath, out string fileName, out string arguments)
        {
            fileName = string.Empty;
            arguments = string.Empty;

            if (!launchPath.StartsWith(StartMenuCommandTokenPrefix, StringComparison.Ordinal))
            {
                return false;
            }

            var payload = launchPath[StartMenuCommandTokenPrefix.Length..];
            var separatorIndex = payload.IndexOf('|');
            if (separatorIndex < 0)
            {
                fileName = payload;
                return !string.IsNullOrWhiteSpace(fileName);
            }

            fileName = payload[..separatorIndex];
            arguments = payload[(separatorIndex + 1)..];
            return !string.IsNullOrWhiteSpace(fileName);
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
                ApplySearchQueryToTextBox();
            }

            ClearStartMenuSearchResultsItems();
            RepositionVisibleStartMenu();
        }

        private void ResetStartMenuSubmenuSessionCache()
        {
            populatedSubmenusThisSession.Clear();
        }

        private void RefreshStartMenuRecentProgramsSection(ContextMenuStrip? targetMenu = null)
        {
            var menu = targetMenu ?? startMenu;
            if (menu == null)
            {
                return;
            }

            foreach (var item in startMenuRecentProgramItems)
            {
                menu.Items.Remove(item);
                item.Dispose();
            }
            startMenuRecentProgramItems.Clear();

            if (startMenuRecentProgramsMaxCount <= 0)
            {
                return;
            }

            var insertIndex = Math.Min(2, menu.Items.Count);
            var recentHeaderItem = CreateMenuItem("Recently used", null, null);
            recentHeaderItem.Enabled = false;
            menu.Items.Insert(insertIndex++, recentHeaderItem);
            startMenuRecentProgramItems.Add(recentHeaderItem);

            var visibleRecentPrograms = startMenuRecentPrograms
                .Where(entry =>
                    entry.IsShellApp ||
                    entry.LaunchPath.StartsWith(StartMenuCommandTokenPrefix, StringComparison.Ordinal) ||
                    File.Exists(entry.LaunchPath) ||
                    Directory.Exists(entry.LaunchPath))
                .Take(startMenuRecentProgramsMaxCount)
                .ToList();

            if (visibleRecentPrograms.Count == 0)
            {
                var emptyItem = CreateMenuItem("(No recently used programs)", null, null);
                emptyItem.Enabled = false;
                menu.Items.Insert(insertIndex++, emptyItem);
                startMenuRecentProgramItems.Add(emptyItem);
            }
            else
            {
                foreach (var entry in visibleRecentPrograms)
                {
                    var localEntry = entry;
                    var recentItem = CreateMenuItem(localEntry.DisplayName, GetProgramSearchEntryIcon(localEntry), (_, _) => LaunchProgramSearchEntry(localEntry, trackRecentUsage: true));
                    AttachProgramSearchContextMenu(recentItem, localEntry, trackRecentUsage: true);
                    menu.Items.Insert(insertIndex++, recentItem);
                    startMenuRecentProgramItems.Add(recentItem);
                }
            }

            var divider = new ToolStripSeparator();
            menu.Items.Insert(insertIndex, divider);
            startMenuRecentProgramItems.Add(divider);
        }

        private void RecordStartMenuRecentProgram(ProgramSearchEntry entry)
        {
            startMenuRecentPrograms.RemoveAll(existing =>
                existing.IsShellApp == entry.IsShellApp &&
                string.Equals(existing.LaunchPath, entry.LaunchPath, StringComparison.OrdinalIgnoreCase));
            startMenuRecentPrograms.Insert(0, entry);

            if (startMenuRecentPrograms.Count > startMenuRecentProgramsMaxCount)
            {
                startMenuRecentPrograms.RemoveRange(startMenuRecentProgramsMaxCount, startMenuRecentPrograms.Count - startMenuRecentProgramsMaxCount);
            }

            RefreshStartMenuRecentProgramsSection();
            SaveTaskbarSettings();
        }

        private void StartMenuSearchTextBox_TextChanged(object? sender, EventArgs e)
        {
            if (suppressStartMenuSearchTextChanged || startMenuSearchTextBox == null)
            {
                return;
            }

            startMenuSearchQuery.Clear();
            startMenuSearchQuery.Append(startMenuSearchTextBox.Text);
            startMenuSearchLastInputUtc = DateTime.UtcNow;
            UpdateStartMenuSearchResultsFromCurrentQuery();
        }

        private void StartMenuSearchTextBox_KeyDown(object? sender, KeyEventArgs e)
        {
            if (!startMenu.Visible)
            {
                return;
            }

            if (!IsStartMenuNavigationVirtualKey((int)e.KeyCode))
            {
                return;
            }

            if (HandleStartMenuSearchKey((int)e.KeyCode))
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }

        private bool IsStartMenuSearchTextBoxFocused()
        {
            return startMenuSearchTextBox?.TextBox.Focused == true;
        }

        private static bool IsStartMenuNavigationVirtualKey(int vkCode)
        {
            return vkCode == (int)Keys.Escape ||
                   vkCode == (int)Keys.Return ||
                   vkCode == (int)Keys.Back ||
                   vkCode == (int)Keys.Tab ||
                   vkCode == (int)Keys.Up ||
                   vkCode == (int)Keys.Down ||
                   vkCode == (int)Keys.Left ||
                   vkCode == (int)Keys.Right ||
                   vkCode == (int)Keys.Home ||
                   vkCode == (int)Keys.End ||
                   vkCode == (int)Keys.PageUp ||
                   vkCode == (int)Keys.PageDown;
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
                    LaunchProgramSearchEntry(selectedEntry, trackRecentUsage: true);
                    startMenu.Close(ToolStripDropDownCloseReason.ItemClicked);
                    return true;
                }

                var topMatch = GetProgramSearchMatches(startMenuSearchQuery.ToString(), 1).FirstOrDefault();

                if (topMatch != null)
                {
                    LaunchProgramSearchEntry(topMatch, trackRecentUsage: true);
                    startMenu.Close(ToolStripDropDownCloseReason.ItemClicked);
                }

                return true;
            }

            if (vkCode == (int)Keys.Down || vkCode == (int)Keys.Up ||
                vkCode == (int)Keys.Home || vkCode == (int)Keys.End ||
                vkCode == (int)Keys.PageUp || vkCode == (int)Keys.PageDown)
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
                var nextIndex = vkCode switch
                {
                    (int)Keys.Down => (currentIndex + 1 + selectableItems.Count) % selectableItems.Count,
                    (int)Keys.Up => (currentIndex - 1 + selectableItems.Count) % selectableItems.Count,
                    (int)Keys.Home => 0,
                    (int)Keys.End => selectableItems.Count - 1,
                    (int)Keys.PageDown => Math.Clamp(currentIndex + 8, 0, selectableItems.Count - 1),
                    (int)Keys.PageUp => Math.Clamp(currentIndex - 8, 0, selectableItems.Count - 1),
                    _ => currentIndex
                };

                if (nextIndex < 0)
                {
                    nextIndex = 0;
                }

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

            if (vkCode == (int)Keys.Home || vkCode == (int)Keys.PageUp)
            {
                navigableItems[0].Select();
                return true;
            }

            if (vkCode == (int)Keys.End || vkCode == (int)Keys.PageDown)
            {
                navigableItems[^1].Select();
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

        private bool TryActivateStartMenuAccelerator(int vkCode)
        {
            if (vkCode < (int)Keys.A || vkCode > (int)Keys.Z)
            {
                return false;
            }

            var accelerator = char.ToUpperInvariant((char)vkCode);
            var candidates = startMenu.Items
                .OfType<ToolStripMenuItem>()
                .Where(item => item.Available && item.Enabled)
                .ToList();

            var currentIndex = candidates.FindIndex(item => item.Selected);
            for (var offset = 1; offset <= candidates.Count; offset++)
            {
                var candidate = candidates[(Math.Max(currentIndex, -1) + offset) % candidates.Count];
                var text = candidate.Text?.TrimStart() ?? string.Empty;
                if (text.Length == 0)
                {
                    continue;
                }

                if (char.ToUpperInvariant(text[0]) != accelerator)
                {
                    continue;
                }

                candidate.Select();
                if (candidate.HasDropDownItems)
                {
                    candidate.ShowDropDown();
                    SelectFirstDropdownItem(candidate);
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

            if (vkCode == (int)Keys.Home)
            {
                dropdownItems[0].Select();
                return true;
            }

            if (vkCode == (int)Keys.End)
            {
                dropdownItems[^1].Select();
                return true;
            }

            if (vkCode == (int)Keys.PageUp || vkCode == (int)Keys.PageDown)
            {
                var pageStep = 8;
                var direction = vkCode == (int)Keys.PageDown ? 1 : -1;
                var targetIndex = Math.Clamp(selectedIndex + (direction * pageStep), 0, dropdownItems.Count - 1);
                dropdownItems[targetIndex].Select();
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

            suppressStartMenuSearchTextChanged = true;
            startMenuSearchTextBox.Text = startMenuSearchQuery.ToString();
            startMenuSearchTextBox.SelectionStart = startMenuSearchTextBox.TextLength;
            suppressStartMenuSearchTextChanged = false;
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
                    var matchItem = CreateMenuItem(match.DisplayName, icon, (_, _) => LaunchProgramSearchEntry(match, trackRecentUsage: true));
                    matchItem.Tag = match;
                    AttachProgramSearchContextMenu(matchItem, match, trackRecentUsage: true);
                    startMenu.Items.Insert(insertIndex + i, matchItem);
                    startMenuSearchResultItems.Add(matchItem);
                }
            }

            RepositionVisibleStartMenu();
        }

        private List<ProgramSearchEntry> GetProgramSearchMatches(string query, int maxCount)
        {
            var uniqueByName = new HashSet<string>(StringComparer.CurrentCultureIgnoreCase);
            var normalizedQuery = query.Trim();

            return GetProgramSearchEntries()
                .Where(entry => entry.DisplayName.Contains(normalizedQuery, StringComparison.OrdinalIgnoreCase))
                .OrderBy(entry =>
                {
                    var name = entry.DisplayName;
                    var startsWith = name.StartsWith(normalizedQuery, StringComparison.OrdinalIgnoreCase);
                    var contains = name.Contains(normalizedQuery, StringComparison.OrdinalIgnoreCase);

                    if (startsWith && !entry.IsShellApp)
                    {
                        return 0;
                    }

                    if (contains && !entry.IsShellApp)
                    {
                        return 1;
                    }

                    if (startsWith)
                    {
                        return 2;
                    }

                    return 3;
                })
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
            lock (programSearchEntriesCacheLock)
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
                        Arguments = "-NoProfile -ExecutionPolicy Bypass -Command \"$OutputEncoding = [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new(); Get-StartApps | ForEach-Object { '{0}|{1}' -f $_.Name, $_.AppID }\"",
                        UseShellExecute = false,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        StandardOutputEncoding = Encoding.UTF8,
                        StandardErrorEncoding = Encoding.UTF8,
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

        private void LaunchProgramSearchEntry(ProgramSearchEntry entry, bool trackRecentUsage = false)
        {
            if (entry.IsShellApp)
            {
                LaunchProcess("explorer.exe", $"shell:AppsFolder\\{entry.LaunchPath}");
                if (trackRecentUsage)
                {
                    RecordStartMenuRecentProgram(entry);
                }
                return;
            }

            if (TryDecodeCommandLaunchPath(entry.LaunchPath, out var fileName, out var arguments))
            {
                if (string.Equals(fileName, "win9xplorer-taskmgr", StringComparison.OrdinalIgnoreCase))
                {
                    OpenWin9xplorerTaskManager();
                }
                else
                {
                    LaunchProcess(fileName, string.IsNullOrWhiteSpace(arguments) ? null : arguments);
                }
            }
            else
            {
                LaunchProcess(entry.LaunchPath);
            }

            if (trackRecentUsage)
            {
                RecordStartMenuRecentProgram(entry);
            }
        }

        private Image? GetProgramSearchEntryIcon(ProgramSearchEntry entry)
        {
            if (!entry.IsShellApp)
            {
                if (entry.LaunchPath.StartsWith(StartMenuCommandTokenPrefix, StringComparison.Ordinal))
                {
                    return LoadMenuAsset("Application Window.png", SystemIcons.Application.ToBitmap());
                }

                if (string.Equals(Path.GetExtension(entry.LaunchPath), ".lnk", StringComparison.OrdinalIgnoreCase))
                {
                    return GetQuickLaunchButtonIcon(entry.LaunchPath);
                }

                return GetSmallShellIcon(entry.LaunchPath, isDirectory: false);
            }

            var shellPath = $"shell:AppsFolder\\{entry.LaunchPath}";
            return GetSmallShellIcon(shellPath, isDirectory: false)
                ?? GetSmallShellIconFromShellPath(shellPath)
                ?? LoadMenuAsset("Application Window.png", SystemIcons.Application.ToBitmap());
        }

        private Image? GetCachedShellIcon(string cacheKey, Func<Image?> loader)
        {
            lock (shellIconCacheLock)
            {
                if (shellIconCache.TryGetValue(cacheKey, out var cached))
                {
                    return (Image)cached.Clone();
                }
            }

            var loaded = loader();
            if (loaded is not Bitmap loadedBitmap)
            {
                loaded?.Dispose();
                return null;
            }

            lock (shellIconCacheLock)
            {
                if (!shellIconCache.ContainsKey(cacheKey))
                {
                    shellIconCache[cacheKey] = (Bitmap)loadedBitmap.Clone();
                }

                return (Image)shellIconCache[cacheKey].Clone();
            }
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

        private Image? GetSmallShellIcon(string pathOrExtension, bool isDirectory)
        {
            var cacheKey = $"shell|{isDirectory}|{pathOrExtension}";
            return GetCachedShellIcon(cacheKey, () =>
            {
                var shfi = new WinApi.SHFILEINFO();
                var flags = WinApi.SHGFI_ICON | WinApi.SHGFI_SMALLICON;
                var isShellNamespacePath = pathOrExtension.StartsWith("shell:", StringComparison.OrdinalIgnoreCase) ||
                                           pathOrExtension.Contains("AppsFolder", StringComparison.OrdinalIgnoreCase);

                if (isDirectory)
                {
                    flags |= WinApi.SHGFI_USEFILEATTRIBUTES;
                    WinApi.SHGetFileInfo(pathOrExtension, WinApi.FILE_ATTRIBUTE_DIRECTORY, ref shfi, (uint)Marshal.SizeOf(shfi), flags);
                }
                else
                {
                    if (!isShellNamespacePath && !File.Exists(pathOrExtension))
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
            });
        }

        private Image? GetSmallShellIconFromShellPath(string shellPath)
        {
            var cacheKey = $"shellpath|{shellPath}";
            return GetCachedShellIcon(cacheKey, () =>
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
            });
        }

        private static Image LoadMenuAsset(string fileName, Image fallback)
        {
            try
            {
                var diskImage = TryLoadAssetFromDisk(fileName);
                if (diskImage != null)
                {
                    return diskImage;
                }

                var assembly = typeof(RetroTaskbarForm).Assembly;
                var resourceName = assembly
                    .GetManifestResourceNames()
                    .FirstOrDefault(name =>
                        name.EndsWith(fileName, StringComparison.OrdinalIgnoreCase));

                if (!string.IsNullOrWhiteSpace(resourceName))
                {
                    using var stream = assembly.GetManifestResourceStream(resourceName);
                    if (stream == null)
                    {
                        return new Bitmap(fallback);
                    }

                    using var image = Image.FromStream(stream);
                    return new Bitmap(image);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to load menu asset '{fileName}': {ex.Message}");
            }

            return new Bitmap(fallback);
        }

        private static Image LoadAssetBySuffix(string fileName, Image fallback)
        {
            try
            {
                var diskImage = TryLoadAssetFromDisk(fileName);
                if (diskImage != null)
                {
                    return diskImage;
                }

                var assembly = typeof(RetroTaskbarForm).Assembly;
                var resourceName = assembly
                    .GetManifestResourceNames()
                    .FirstOrDefault(name => name.EndsWith(fileName, StringComparison.OrdinalIgnoreCase));

                if (!string.IsNullOrWhiteSpace(resourceName))
                {
                    using var stream = assembly.GetManifestResourceStream(resourceName);
                    if (stream != null)
                    {
                        using var image = Image.FromStream(stream);
                        return new Bitmap(image);
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to load asset '{fileName}': {ex.Message}");
            }

            return new Bitmap(fallback);
        }

        private static Image? TryLoadAssetFromDisk(string fileName)
        {
            var candidates = new[]
            {
                Path.Combine(AppContext.BaseDirectory, "Windows XP Icons", fileName),
                Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Windows XP Icons", fileName)),
                Path.Combine(Environment.CurrentDirectory, "Windows XP Icons", fileName)
            };

            foreach (var path in candidates)
            {
                try
                {
                    if (!File.Exists(path))
                    {
                        continue;
                    }

                    using var stream = File.OpenRead(path);
                    using var image = Image.FromStream(stream);
                    return new Bitmap(image);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Failed to load asset from '{path}': {ex.Message}");
                }
            }

            return null;
        }

        private static Image LoadVolumeTrayIconAsset(Image fallback)
        {
            var candidates = new[]
            {
                Path.Combine(AppContext.BaseDirectory, "Windows XP Icons", "Volume.png"),
                Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Windows XP Icons", "Volume.png")),
                Path.Combine(Environment.CurrentDirectory, "Windows XP Icons", "Volume.png")
            };

            foreach (var path in candidates)
            {
                try
                {
                    if (!File.Exists(path))
                    {
                        continue;
                    }

                    using var stream = File.OpenRead(path);
                    using var image = Image.FromStream(stream);
                    return new Bitmap(image);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Failed to load volume icon from '{path}': {ex.Message}");
                }
            }

            return new Bitmap(fallback);
        }

        private static Image LoadPlayerControlIconAsset(string fileName, Image fallback)
        {
            var candidates = new[]
            {
                Path.Combine(AppContext.BaseDirectory, "player Icons", fileName),
                Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "player Icons", fileName)),
                Path.Combine(Environment.CurrentDirectory, "player Icons", fileName)
            };

            foreach (var path in candidates)
            {
                try
                {
                    if (!File.Exists(path))
                    {
                        continue;
                    }

                    using var stream = File.OpenRead(path);
                    using var image = Image.FromStream(stream);
                    return new Bitmap(image);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Failed to load player icon from '{path}': {ex.Message}");
                }
            }

            return new Bitmap(fallback);
        }

        private ContextMenuStrip BuildTaskWindowMenu()
        {
            var menu = new ContextMenuStrip
            {
                ShowImageMargin = false,
                ShowCheckMargin = false,
                Font = normalFont
            };
            ApplyWin9xContextMenuStyle(menu);

            var restoreItem = new ToolStripMenuItem("Restore", null, (_, _) => InvokeWindowAction(window => window.Restore()));
            var minimizeItem = new ToolStripMenuItem("Minimize", null, (_, _) => InvokeWindowAction(window => window.Minimize()));
            var maximizeItem = new ToolStripMenuItem("Maximize", null, (_, _) => InvokeWindowAction(window => window.Maximize()));
            var bringToFrontItem = new ToolStripMenuItem("Bring To Front", null, (_, _) => InvokeWindowAction(window => window.BringToFront()));
            var openLocationItem = new ToolStripMenuItem("Open File Location", null, (_, _) => OpenContextWindowProcessLocation());
            var copyInfoItem = new ToolStripMenuItem("Copy Window Info", null, (_, _) => CopyContextWindowInfo());
            var endProcessItem = new ToolStripMenuItem("End Process", null, (_, _) => EndContextWindowProcess());
            var closeItem = new ToolStripMenuItem("Close", null, (_, _) => InvokeWindowAction(window => window.Close()));

            menu.Items.Add(restoreItem);
            menu.Items.Add(minimizeItem);
            menu.Items.Add(maximizeItem);
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add(bringToFrontItem);
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add(openLocationItem);
            menu.Items.Add(copyInfoItem);
            menu.Items.Add(endProcessItem);
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add(closeItem);

            menu.Opening += (_, _) =>
            {
                var hasWindow = windowsByHandle.TryGetValue(contextMenuWindowHandle, out var window);
                var processId = hasWindow ? GetProcessIdForWindow(contextMenuWindowHandle) : 0;
                var processPath = processId > 0 ? GetProcessPath(processId) : string.Empty;
                restoreItem.Enabled = hasWindow && window!.IsMinimized;
                minimizeItem.Enabled = hasWindow && !window!.IsMinimized;
                maximizeItem.Enabled = hasWindow;
                bringToFrontItem.Enabled = hasWindow;
                openLocationItem.Enabled = hasWindow && !string.IsNullOrWhiteSpace(processPath) && File.Exists(processPath);
                copyInfoItem.Enabled = hasWindow;
                endProcessItem.Enabled = hasWindow && processId > 0 && processId != Environment.ProcessId;
                closeItem.Enabled = hasWindow;
            };

            return menu;
        }

        private ContextMenuStrip BuildTaskbarContextMenu()
        {
            var menu = new ContextMenuStrip
            {
                ShowCheckMargin = true,
                ShowImageMargin = false,
                Font = normalFont
            };
            ApplyWin9xContextMenuStyle(menu);

            var toolbarsItem = new ToolStripMenuItem("Toolbars");
            var quickLaunchToolbarItem = new ToolStripMenuItem("Quick Launch", null, (_, _) => ToggleQuickLaunchToolbar());
            var addressToolbarItem = new ToolStripMenuItem("Address", null, (_, _) => ToggleAddressToolbar());
            var volumeToolbarItem = new ToolStripMenuItem("Volume", null, (_, _) => ToggleVolumeToolbar());
            var spotifyToolbarItem = new ToolStripMenuItem("Spotify Controls", null, (_, _) => ToggleSpotifyToolbar());
            var linksToolbarItem = new ToolStripMenuItem("Links") { Enabled = false };
            var desktopToolbarItem = new ToolStripMenuItem("Desktop") { Enabled = false };
            var addressLeftItem = new ToolStripMenuItem("Address Left of Quick Launch", null, (_, _) => SetToolbarOrder(addressFirst: true));
            var quickLaunchLeftItem = new ToolStripMenuItem("Quick Launch Left of Address", null, (_, _) => SetToolbarOrder(addressFirst: false));
            toolbarsItem.DropDownItems.Add(quickLaunchToolbarItem);
            toolbarsItem.DropDownItems.Add(addressToolbarItem);
            toolbarsItem.DropDownItems.Add(volumeToolbarItem);
            toolbarsItem.DropDownItems.Add(spotifyToolbarItem);
            toolbarsItem.DropDownItems.Add(linksToolbarItem);
            toolbarsItem.DropDownItems.Add(desktopToolbarItem);
            toolbarsItem.DropDownItems.Add(new ToolStripSeparator());
            toolbarsItem.DropDownItems.Add(addressLeftItem);
            toolbarsItem.DropDownItems.Add(quickLaunchLeftItem);

            menu.Items.Add(toolbarsItem);
            menu.Items.Add(new ToolStripSeparator());
            var appMenuItem = new ToolStripMenuItem("win9xplorer");
            appMenuItem.DropDownItems.Add("Diagnostics...", null, (_, _) => ShowDiagnosticsPanel());
            appMenuItem.DropDownItems.Add("Options...", null, (_, _) => ShowOptionsDialog());
            appMenuItem.DropDownItems.Add(new ToolStripSeparator());
            appMenuItem.DropDownItems.Add("Exit App", null, (_, _) => RequestAppExit());
            menu.Items.Add(appMenuItem);
            menu.Items.Add(new ToolStripSeparator());
            var cascadeWindowsItem = new ToolStripMenuItem("Cascade Windows", null, (_, _) => CascadeWindows());
            var tileHorizontallyItem = new ToolStripMenuItem("Tile Windows Horizontally", null, (_, _) => TileWindowsHorizontally());
            var tileVerticallyItem = new ToolStripMenuItem("Tile Windows Vertically", null, (_, _) => TileWindowsVertically());
            var tileGridItem = new ToolStripMenuItem("Tile Windows in Grids", null, (_, _) => TileWindowsInGrid());
            var restoreMinimizedItem = new ToolStripMenuItem("Restore Minimized Windows", null, (_, _) => RestoreMinimizedWindows());
            menu.Items.Add(cascadeWindowsItem);
            menu.Items.Add(tileHorizontallyItem);
            menu.Items.Add(tileVerticallyItem);
            menu.Items.Add(tileGridItem);
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("Show the Desktop", null, (_, _) => ShowDesktop());
            menu.Items.Add(restoreMinimizedItem);
            menu.Items.Add("Task Manager", null, (_, _) => OpenTaskManager());
            menu.Items.Add("win9xplorer Task Manager", null, (_, _) => OpenWin9xplorerTaskManager());
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("Run...", null, (_, _) => LaunchProcess("rundll32.exe", "shell32.dll,#61"));
            menu.Items.Add("File Explorer", null, (_, _) => BringExplorerToFront());
            menu.Items.Add("Command Prompt", null, (_, _) => LaunchProcess("cmd.exe"));
            menu.Items.Add("PowerShell", null, (_, _) => LaunchProcess("powershell.exe"));
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("Refresh Quick Launch", null, (_, _) => LoadQuickLaunchButtons());
            menu.Items.Add(new ToolStripSeparator());
            var rowsItem = new ToolStripMenuItem("Rows");
            var oneRowItem = new ToolStripMenuItem("1 Row", null, (_, _) => SetTaskbarRows(1));
            var twoRowsItem = new ToolStripMenuItem("2 Rows", null, (_, _) => SetTaskbarRows(2));
            var threeRowsItem = new ToolStripMenuItem("3 Rows", null, (_, _) => SetTaskbarRows(3));
            rowsItem.DropDownItems.Add(oneRowItem);
            rowsItem.DropDownItems.Add(twoRowsItem);
            rowsItem.DropDownItems.Add(threeRowsItem);
            menu.Items.Add(rowsItem);
            var autoHideItem = new ToolStripMenuItem("Auto-hide Taskbar")
            {
                CheckOnClick = true
            };
            autoHideItem.CheckedChanged += (_, _) => SetAutoHideTaskbar(autoHideItem.Checked);
            menu.Items.Add(autoHideItem);
            var lockItem = new ToolStripMenuItem("Lock Taskbar")
            {
                CheckOnClick = true
            };
            lockItem.CheckedChanged += (_, _) => SetTaskbarLocked(lockItem.Checked);
            menu.Items.Add(lockItem);
            var topLevelSubmenuItems = new[] { toolbarsItem, appMenuItem, rowsItem };
            foreach (var parentItem in topLevelSubmenuItems)
            {
                parentItem.MouseEnter += (_, _) => BeginInvoke(new Action(() => ShowTaskbarTopLevelSubmenu(menu, parentItem)));
                parentItem.MouseHover += (_, _) => BeginInvoke(new Action(() => ShowTaskbarTopLevelSubmenu(menu, parentItem)));
                parentItem.MouseMove += (_, _) => BeginInvoke(new Action(() => ShowTaskbarTopLevelSubmenu(menu, parentItem)));
            }

            menu.MouseMove += (_, e) =>
            {
                if (menu.GetItemAt(e.Location) is ToolStripMenuItem hoverItem && hoverItem.HasDropDownItems)
                {
                    BeginInvoke(new Action(() => ShowTaskbarTopLevelSubmenu(menu, hoverItem)));
                }
            };

            menu.Opening += (_, _) =>
            {
                quickLaunchToolbarItem.Checked = showQuickLaunchToolbar;
                addressToolbarItem.Checked = showAddressToolbar;
                volumeToolbarItem.Checked = showVolumeToolbar;
                spotifyToolbarItem.Checked = showSpotifyToolbar;
                var canOrderLeftRail = showQuickLaunchToolbar && showAddressToolbar && !ShouldUseDedicatedAddressToolbarRow();
                addressLeftItem.Enabled = canOrderLeftRail;
                quickLaunchLeftItem.Enabled = canOrderLeftRail;
                addressLeftItem.Checked = canOrderLeftRail && addressToolbarBeforeQuickLaunch;
                quickLaunchLeftItem.Checked = canOrderLeftRail && !addressToolbarBeforeQuickLaunch;
                var arrangeCount = GetArrangeableWindowHandles().Count;
                cascadeWindowsItem.Enabled = arrangeCount > 0;
                tileHorizontallyItem.Enabled = arrangeCount > 0;
                tileVerticallyItem.Enabled = arrangeCount > 0;
                tileGridItem.Enabled = arrangeCount > 0;
                restoreMinimizedItem.Enabled = windowsByHandle.Values.Any(window => window.IsMinimized);
                oneRowItem.Checked = taskbarRows == 1;
                twoRowsItem.Checked = taskbarRows == 2;
                threeRowsItem.Checked = taskbarRows == 3;
                rowsItem.Enabled = true;
                autoHideItem.Checked = autoHideTaskbar;
                lockItem.Checked = taskbarLocked;
            };
            return menu;
        }

        private void PopulateSystemToolsMenu(ToolStripItemCollection items)
        {
            items.Add(CreateMenuItem("Control Panel", SystemIcons.Shield.ToBitmap(), (_, _) => LaunchProcess("control.exe"), new ProgramSearchEntry("Control Panel", EncodeCommandLaunchPath("control.exe", null), IsShellApp: false)));
            items.Add(new ToolStripSeparator());
            items.Add(CreateMenuItem("Task Manager", LoadMenuAsset("Task Manager.png", SystemIcons.Application.ToBitmap()), (_, _) => OpenTaskManager(), new ProgramSearchEntry("Task Manager", EncodeCommandLaunchPath("taskmgr.exe", null), IsShellApp: false)));
            items.Add(CreateMenuItem("win9xplorer Task Manager", LoadMenuAsset("Task Manager.png", SystemIcons.Application.ToBitmap()), (_, _) => OpenWin9xplorerTaskManager(), new ProgramSearchEntry("win9xplorer Task Manager", EncodeCommandLaunchPath("win9xplorer-taskmgr", null), IsShellApp: false)));
            items.Add(new ToolStripSeparator());
            items.Add(CreateMenuItem("Command Prompt", LoadMenuAsset("Command Prompt.png", SystemIcons.Application.ToBitmap()), (_, _) => LaunchProcess("cmd.exe"), new ProgramSearchEntry("Command Prompt", EncodeCommandLaunchPath("cmd.exe", null), IsShellApp: false)));
            items.Add(CreateMenuItem("PowerShell", SystemIcons.Application.ToBitmap(), (_, _) => LaunchProcess("powershell.exe"), new ProgramSearchEntry("PowerShell", EncodeCommandLaunchPath("powershell.exe", null), IsShellApp: false)));
            items.Add(CreateMenuItem("Environment Variables...", settingsMenuIcon, (_, _) => OpenEnvironmentVariablesSettings(), new ProgramSearchEntry("Environment Variables", EncodeCommandLaunchPath("rundll32.exe", "sysdm.cpl,EditEnvironmentVariables"), IsShellApp: false)));
            items.Add(CreateMenuItem("Run...", runMenuIcon, (_, _) => LaunchProcess("rundll32.exe", "shell32.dll,#61"), new ProgramSearchEntry("Run...", EncodeCommandLaunchPath("rundll32.exe", "shell32.dll,#61"), IsShellApp: false)));
            items.Add(new ToolStripSeparator());
            items.Add(CreateMenuItem("Computer Management", LoadMenuAsset("Computer Management.png", SystemIcons.Application.ToBitmap()), (_, _) => LaunchProcess("compmgmt.msc"), new ProgramSearchEntry("Computer Management", EncodeCommandLaunchPath("compmgmt.msc", null), IsShellApp: false)));
            items.Add(CreateMenuItem("Device Manager", LoadMenuAsset("Chip.png", SystemIcons.Application.ToBitmap()), (_, _) => LaunchProcess("devmgmt.msc"), new ProgramSearchEntry("Device Manager", EncodeCommandLaunchPath("devmgmt.msc", null), IsShellApp: false)));
            items.Add(CreateMenuItem("Disk Management", LoadMenuAsset("Disk Management.png", SystemIcons.Application.ToBitmap()), (_, _) => LaunchProcess("diskmgmt.msc"), new ProgramSearchEntry("Disk Management", EncodeCommandLaunchPath("diskmgmt.msc", null), IsShellApp: false)));
            items.Add(CreateMenuItem("Event Viewer", LoadMenuAsset("Event Viewer.png", SystemIcons.Application.ToBitmap()), (_, _) => LaunchProcess("eventvwr.msc"), new ProgramSearchEntry("Event Viewer", EncodeCommandLaunchPath("eventvwr.msc", null), IsShellApp: false)));
            items.Add(CreateMenuItem("Services", LoadMenuAsset("Services.png", SystemIcons.Application.ToBitmap()), (_, _) => LaunchProcess("services.msc"), new ProgramSearchEntry("Services", EncodeCommandLaunchPath("services.msc", null), IsShellApp: false)));
            items.Add(CreateMenuItem("Registry Editor", LoadMenuAsset("Registry Editor.png", SystemIcons.Application.ToBitmap()), (_, _) => LaunchProcess("regedit.exe"), new ProgramSearchEntry("Registry Editor", EncodeCommandLaunchPath("regedit.exe", null), IsShellApp: false)));
            items.Add(CreateMenuItem("System Information", LoadMenuAsset("Information.png", SystemIcons.Information.ToBitmap()), (_, _) => LaunchProcess("msinfo32.exe"), new ProgramSearchEntry("System Information", EncodeCommandLaunchPath("msinfo32.exe", null), IsShellApp: false)));
        }

        private static void ShowTaskbarTopLevelSubmenu(ContextMenuStrip menu, ToolStripMenuItem item)
        {
            if (!menu.Visible || !item.HasDropDownItems || !item.Available || !item.Enabled)
            {
                return;
            }

            foreach (var sibling in menu.Items.OfType<ToolStripMenuItem>())
            {
                if (!ReferenceEquals(sibling, item) && sibling.DropDown.Visible)
                {
                    sibling.HideDropDown();
                }
            }

            item.Select();
            if (!item.DropDown.Visible)
            {
                item.ShowDropDown();
            }
        }

        private void ToggleQuickLaunchToolbar()
        {
            showQuickLaunchToolbar = !showQuickLaunchToolbar;
            ApplyToolbarLayout();
            SaveTaskbarSettings();
        }

        private void ToggleAddressToolbar()
        {
            showAddressToolbar = !showAddressToolbar;
            ApplyToolbarLayout();
            SaveTaskbarSettings();
        }

        private void ToggleSpotifyToolbar()
        {
            showSpotifyToolbar = !showSpotifyToolbar;
            ApplyToolbarLayout();
            SaveTaskbarSettings();
        }

        private void ToggleVolumeToolbar()
        {
            showVolumeToolbar = !showVolumeToolbar;
            ApplyToolbarLayout();
            SaveTaskbarSettings();
        }

        private void ExecuteSpotifyCommand(SpotifyController.SpotifyCommand command)
        {
            if (spotifyController.TryExecute(command))
            {
                return;
            }

            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "spotify:",
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Unable to launch Spotify: {ex.Message}");
            }
        }

        private void SetToolbarOrder(bool addressFirst)
        {
            addressToolbarBeforeQuickLaunch = addressFirst;
            ApplyToolbarLayout();
            SaveTaskbarSettings();
        }

        private void SetAutoHideTaskbar(bool enabled)
        {
            if (autoHideTaskbar == enabled)
            {
                return;
            }

            autoHideTaskbar = enabled;
            SetTaskbarBounds();
            SaveTaskbarSettings();
        }

        private void SetTaskbarRows(int rows)
        {
            var newRows = Math.Clamp(rows, 1, 3);
            if (taskbarRows == newRows)
            {
                return;
            }

            taskbarRows = newRows;
            SetTaskbarBounds();
            RefreshTaskbar();
            ResizeTrayPanel();
            RefreshTrayHostPlacement();
            SaveTaskbarSettings();
        }

        private int GetTaskButtonRows()
        {
            var rows = taskbarRows;
            if (ShouldUseDedicatedAddressToolbarRow())
            {
                rows--;
            }

            return Math.Clamp(rows, 1, 3);
        }

        private void ToolbarGripPanel_Paint(object? sender, PaintEventArgs e)
        {
            if (sender is not Panel panel)
            {
                return;
            }

            e.Graphics.Clear(taskbarBaseColor);
            var left = 2;
            var top = 2;
            var bottom = Math.Max(top, panel.Height - 5);
            DrawToolbarGrip(e.Graphics, left, top, bottom);
        }

        private void DrawToolbarGrip(Graphics graphics, int left, int top, int bottom)
        {
            using var lightBrush = new SolidBrush(taskbarLightColor);
            using var baseBrush = new SolidBrush(taskbarBaseColor);
            using var darkBrush = new SolidBrush(taskbarDarkColor);

            for (var y = top; y <= bottom; y++)
            {
                var firstBrush = y == bottom ? darkBrush : lightBrush;
                var secondBrush = y == top
                    ? lightBrush
                    : y == bottom
                        ? darkBrush
                        : baseBrush;

                graphics.FillRectangle(firstBrush, left, y, 1, 1);
                graphics.FillRectangle(secondBrush, left + 1, y, 1, 1);
                graphics.FillRectangle(darkBrush, left + 2, y, 1, 1);
            }
        }

        private void ToolbarSeparatorPanel_Paint(object? sender, PaintEventArgs e)
        {
            if (sender is not Panel panel)
            {
                return;
            }

            e.Graphics.Clear(taskbarBaseColor);
            var center = panel.Width / 2;
            var top = 0;
            var bottom = Math.Max(top, panel.Height - 3);
            using var darkPen = new Pen(taskbarDarkColor);
            using var lightPen = new Pen(taskbarLightColor);
            e.Graphics.DrawLine(darkPen, center - 1, top, center - 1, bottom);
            e.Graphics.DrawLine(lightPen, center, top, center, bottom);
        }

        private void AddressRowSeparatorPanel_Paint(object? sender, PaintEventArgs e)
        {
            if (sender is not Panel panel)
            {
                return;
            }

            e.Graphics.Clear(taskbarBaseColor);
            var middle = panel.Height / 2;
            var yDark = Math.Max(0, middle - 1);
            var yLight = Math.Min(panel.Height - 1, middle);
            var startX = 0;
            var endX = Math.Max(startX, panel.Width - 1);
            using var darkPen = new Pen(taskbarDarkColor);
            using var lightPen = new Pen(taskbarLightColor);
            e.Graphics.DrawLine(darkPen, startX, yDark, endX, yDark);
            e.Graphics.DrawLine(lightPen, startX, yLight, endX, yLight);
        }

        private void ToolbarGripPanel_MouseDown(object? sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left || taskbarLocked)
            {
                return;
            }

            if (sender is not Panel panel || panel.Tag is not string toolbarId)
            {
                return;
            }

            if (!showQuickLaunchToolbar || !showAddressToolbar || ShouldUseDedicatedAddressToolbarRow())
            {
                return;
            }

            isToolbarReorderDragging = true;
            toolbarOrderChangedDuringDrag = false;
            draggingToolbarId = toolbarId;
            panel.Capture = true;
        }

        private void ToolbarGripPanel_MouseMove(object? sender, MouseEventArgs e)
        {
            if (!isToolbarReorderDragging || string.IsNullOrWhiteSpace(draggingToolbarId))
            {
                return;
            }

            var cursorPoint = contentPanel.PointToClient(Cursor.Position);
            if (draggingToolbarId.Equals("Address", StringComparison.OrdinalIgnoreCase))
            {
                var quickLaunchMidX = quickLaunchPanel.Left + (quickLaunchPanel.Width / 2);
                var shouldBeAddressFirst = cursorPoint.X < quickLaunchMidX;
                if (shouldBeAddressFirst != addressToolbarBeforeQuickLaunch)
                {
                    addressToolbarBeforeQuickLaunch = shouldBeAddressFirst;
                    toolbarOrderChangedDuringDrag = true;
                    ApplyToolbarLayout(refreshBounds: false);
                }
            }
            else
            {
                var addressMidX = addressToolbarPanel.Left + (addressToolbarPanel.Width / 2);
                var shouldBeAddressFirst = cursorPoint.X > addressMidX;
                if (shouldBeAddressFirst != addressToolbarBeforeQuickLaunch)
                {
                    addressToolbarBeforeQuickLaunch = shouldBeAddressFirst;
                    toolbarOrderChangedDuringDrag = true;
                    ApplyToolbarLayout(refreshBounds: false);
                }
            }
        }

        private void ToolbarGripPanel_MouseUp(object? sender, MouseEventArgs e)
        {
            if (!isToolbarReorderDragging)
            {
                return;
            }

            isToolbarReorderDragging = false;
            draggingToolbarId = null;
            if (sender is Panel panel)
            {
                panel.Capture = false;
            }

            if (toolbarOrderChangedDuringDrag)
            {
                toolbarOrderChangedDuringDrag = false;
                ApplyToolbarLayout();
                SaveTaskbarSettings();
            }
        }

        private void OpenTaskManager()
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "taskmgr.exe",
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to launch Task Manager: {ex.Message}");
            }
        }

        private void OpenWin9xplorerTaskManager()
        {
            if (taskManagerForm != null && !taskManagerForm.IsDisposed)
            {
                taskManagerForm.Show();
                taskManagerForm.BringToFront();
                taskManagerForm.Activate();
                return;
            }

            try
            {
                taskManagerForm = new TaskManagerForm();
                taskManagerForm.FormClosed += (_, _) => taskManagerForm = null;
                taskManagerForm.Show(this);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to show Task Manager: {ex.Message}");
            }
        }

        private List<IntPtr> GetArrangeableWindowHandles()
        {
            return windowsByHandle.Keys
                .Where(handle => handle != IntPtr.Zero && WinApi.GetWindow(handle, WinApi.GW_OWNER) == IntPtr.Zero)
                .ToList();
        }

        private void CascadeWindows()
        {
            var handles = GetArrangeableWindowHandles();
            if (handles.Count == 0)
            {
                return;
            }

            var area = Screen.PrimaryScreen?.WorkingArea ?? new Rectangle(0, 0, 1024, 768);
            var offset = Math.Max(20, Math.Min(36, area.Height / 18));
            var width = Math.Max(280, area.Width - (offset * Math.Max(0, handles.Count - 1)));
            var height = Math.Max(180, area.Height - (offset * Math.Max(0, handles.Count - 1)));

            for (var i = 0; i < handles.Count; i++)
            {
                var handle = handles[i];
                WinApi.ShowWindow(handle, WinApi.SW_RESTORE);
                var x = Math.Min(area.Left + (offset * i), Math.Max(area.Left, area.Right - width));
                var y = Math.Min(area.Top + (offset * i), Math.Max(area.Top, area.Bottom - height));
                WinApi.MoveWindow(handle, x, y, width, height, true);
            }
        }

        private void TileWindowsHorizontally()
        {
            var handles = GetArrangeableWindowHandles();
            if (handles.Count == 0)
            {
                return;
            }

            var area = Screen.PrimaryScreen?.WorkingArea ?? new Rectangle(0, 0, 1024, 768);
            var rows = handles.Count;
            var rowHeight = Math.Max(120, area.Height / rows);

            for (var i = 0; i < handles.Count; i++)
            {
                var handle = handles[i];
                WinApi.ShowWindow(handle, WinApi.SW_RESTORE);
                var y = area.Top + (rowHeight * i);
                var height = (i == handles.Count - 1) ? area.Bottom - y : rowHeight;
                WinApi.MoveWindow(handle, area.Left, y, area.Width, height, true);
            }
        }

        private void TileWindowsVertically()
        {
            var handles = GetArrangeableWindowHandles();
            if (handles.Count == 0)
            {
                return;
            }

            var area = Screen.PrimaryScreen?.WorkingArea ?? new Rectangle(0, 0, 1024, 768);
            var columns = handles.Count;
            var columnWidth = Math.Max(160, area.Width / columns);

            for (var i = 0; i < handles.Count; i++)
            {
                var handle = handles[i];
                WinApi.ShowWindow(handle, WinApi.SW_RESTORE);
                var x = area.Left + (columnWidth * i);
                var width = (i == handles.Count - 1) ? area.Right - x : columnWidth;
                WinApi.MoveWindow(handle, x, area.Top, width, area.Height, true);
            }
        }

        private void TileWindowsInGrid()
        {
            var handles = GetArrangeableWindowHandles();
            if (handles.Count == 0)
            {
                return;
            }

            var area = Screen.PrimaryScreen?.WorkingArea ?? new Rectangle(0, 0, 1024, 768);
            var count = handles.Count;
            var columns = Math.Max(1, (int)Math.Ceiling(Math.Sqrt(count)));
            var rows = Math.Max(1, (int)Math.Ceiling(count / (double)columns));
            var cellWidth = Math.Max(160, area.Width / columns);
            var cellHeight = Math.Max(120, area.Height / rows);

            for (var i = 0; i < count; i++)
            {
                var handle = handles[i];
                WinApi.ShowWindow(handle, WinApi.SW_RESTORE);

                var row = i / columns;
                var col = i % columns;
                var x = area.Left + (col * cellWidth);
                var y = area.Top + (row * cellHeight);

                var width = (col == columns - 1) ? area.Right - x : cellWidth;
                var height = (row == rows - 1) ? area.Bottom - y : cellHeight;

                WinApi.MoveWindow(handle, x, y, width, height, true);
            }
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

        private ContextMenuStrip BuildLanguageIndicatorContextMenu()
        {
            var menu = new ContextMenuStrip
            {
                ShowImageMargin = false,
                Font = normalFont
            };
            ApplyWin9xContextMenuStyle(menu);

            menu.Opening += (_, _) => PopulateLanguageIndicatorContextMenu(menu);
            return menu;
        }

        private ContextMenuStrip BuildImeIndicatorContextMenu()
        {
            var menu = new ContextMenuStrip
            {
                ShowImageMargin = false,
                Font = normalFont
            };
            ApplyWin9xContextMenuStyle(menu);

            menu.Opening += (_, _) => PopulateImeIndicatorContextMenu(menu);

            return menu;
        }

        private void PopulateImeIndicatorContextMenu(ContextMenuStrip menu)
        {
            menu.Items.Clear();

            var imeTargetWindow = GetImeTargetWindowHandle();
            if (PopulateNativeImeMenu(menu, imeTargetWindow))
            {
                return;
            }

            if (!TryGetImeStatus(imeTargetWindow, out var imeOpen, out var conversionMode, out var sentenceMode))
            {
                menu.Items.Add("IME options unavailable").Enabled = false;
                return;
            }

            var nativeInputItem = new ToolStripMenuItem("Native input", null, (_, _) =>
            {
                ApplyImeStateToForegroundWindow(true, conversionMode | WinApi.IME_CMODE_NATIVE, sentenceMode);
            })
            {
                Checked = imeOpen
            };

            var alphanumericItem = new ToolStripMenuItem("Alphanumeric", null, (_, _) =>
            {
                ApplyImeStateToForegroundWindow(false, conversionMode & ~WinApi.IME_CMODE_NATIVE, sentenceMode);
            })
            {
                Checked = !imeOpen
            };

            menu.Items.Add(nativeInputItem);
            menu.Items.Add(alphanumericItem);
        }

        private bool PopulateNativeImeMenu(ContextMenuStrip menu, IntPtr targetWindow)
        {
            if (targetWindow == IntPtr.Zero)
            {
                return false;
            }

            var inputContext = WinApi.ImmGetContext(targetWindow);
            if (inputContext == IntPtr.Zero)
            {
                return false;
            }

            try
            {
                var allMenuTypes = WinApi.IGIMII_CMODE |
                                   WinApi.IGIMII_SMODE |
                                   WinApi.IGIMII_CONFIGURE |
                                   WinApi.IGIMII_TOOLS |
                                   WinApi.IGIMII_HELP |
                                   WinApi.IGIMII_OTHER |
                                   WinApi.IGIMII_INPUTTOOLS;
                return AppendNativeImeMenuItems(menu.Items, inputContext, allMenuTypes, parentMenuItem: null) > 0;
            }
            finally
            {
                WinApi.ImmReleaseContext(targetWindow, inputContext);
            }
        }

        private int AppendNativeImeMenuItems(ToolStripItemCollection targetItems, IntPtr inputContext, uint menuType, WinApi.IMEMENUITEMINFOW? parentMenuItem)
        {
            uint entryCount;
            if (parentMenuItem.HasValue)
            {
                var parent = parentMenuItem.Value;
                parent.cbSize = (uint)Marshal.SizeOf<WinApi.IMEMENUITEMINFOW>();
                entryCount = WinApi.ImmGetImeMenuItemsW(inputContext, WinApi.IGIMIF_RIGHTMENU, menuType, ref parent, null, 0);
            }
            else
            {
                entryCount = WinApi.ImmGetImeMenuItemsW(inputContext, WinApi.IGIMIF_RIGHTMENU, menuType, IntPtr.Zero, null, 0);
            }

            if (entryCount == 0)
            {
                return 0;
            }

            var itemsBuffer = new WinApi.IMEMENUITEMINFOW[entryCount];
            var entrySize = (uint)Marshal.SizeOf<WinApi.IMEMENUITEMINFOW>();
            for (var i = 0; i < itemsBuffer.Length; i++)
            {
                itemsBuffer[i].cbSize = entrySize;
                itemsBuffer[i].szString = string.Empty;
            }

            uint copiedCount;
            if (parentMenuItem.HasValue)
            {
                var parent = parentMenuItem.Value;
                parent.cbSize = entrySize;
                copiedCount = WinApi.ImmGetImeMenuItemsW(inputContext, WinApi.IGIMIF_RIGHTMENU, menuType, ref parent, itemsBuffer, (uint)(itemsBuffer.Length * entrySize));
            }
            else
            {
                copiedCount = WinApi.ImmGetImeMenuItemsW(inputContext, WinApi.IGIMIF_RIGHTMENU, menuType, IntPtr.Zero, itemsBuffer, (uint)(itemsBuffer.Length * entrySize));
            }

            var addedCount = 0;
            for (var i = 0; i < copiedCount && i < itemsBuffer.Length; i++)
            {
                var item = itemsBuffer[i];
                if ((item.fType & WinApi.IMFT_SEPARATOR) != 0)
                {
                    targetItems.Add(new ToolStripSeparator());
                    addedCount++;
                    continue;
                }

                var text = string.IsNullOrWhiteSpace(item.szString) ? "(IME)" : item.szString;
                var menuItem = new ToolStripMenuItem(text)
                {
                    Checked = (item.fState & WinApi.IMFS_CHECKED) != 0,
                    Enabled = (item.fState & WinApi.IMFS_DISABLED) == 0,
                    Font = (item.fState & WinApi.IMFS_DEFAULT) != 0 ? new Font(normalFont, FontStyle.Bold) : normalFont
                };

                var commandId = item.wID;
                var itemData = item.dwItemData;
                menuItem.Click += (_, _) => InvokeNativeImeMenuCommand(commandId, itemData);

                if ((item.fType & WinApi.IMFT_SUBMENU) != 0)
                {
                    AppendNativeImeMenuItems(menuItem.DropDownItems, inputContext, menuType, item);
                }

                targetItems.Add(menuItem);
                addedCount++;
            }

            return addedCount;
        }

        private void InvokeNativeImeMenuCommand(uint commandId, uint itemData)
        {
            var imeTargetWindow = GetImeTargetWindowHandle();
            if (imeTargetWindow == IntPtr.Zero)
            {
                return;
            }

            var inputContext = WinApi.ImmGetContext(imeTargetWindow);
            if (inputContext == IntPtr.Zero)
            {
                return;
            }

            try
            {
                WinApi.ImmNotifyIME(inputContext, WinApi.NI_IMEMENUSELECTED, commandId, itemData);
            }
            finally
            {
                WinApi.ImmReleaseContext(imeTargetWindow, inputContext);
            }

            UpdateInputMethodIndicators();
        }

        private void PopulateLanguageIndicatorContextMenu(ContextMenuStrip menu)
        {
            menu.Items.Clear();

            var installedLanguages = InputLanguage.InstalledInputLanguages.Cast<InputLanguage>().ToList();
            var currentLanguage = InputLanguage.CurrentInputLanguage;

            foreach (var inputLanguage in installedLanguages)
            {
                var culture = inputLanguage.Culture;
                var languageCode = GetLanguageIndicatorText(culture);
                var menuText = $"{languageCode}  {culture.NativeName}";
                var item = new ToolStripMenuItem(menuText, null, (_, _) =>
                {
                    InputLanguage.CurrentInputLanguage = inputLanguage;
                    UpdateInputMethodIndicators();
                })
                {
                    Checked = currentLanguage != null && inputLanguage.LayoutName.Equals(currentLanguage.LayoutName, StringComparison.OrdinalIgnoreCase)
                };

                menu.Items.Add(item);
            }

            if (menu.Items.Count > 0)
            {
                menu.Items.Add(new ToolStripSeparator());
            }

            menu.Items.Add("Windows IME Settings...", null, (_, _) => OpenWindowsImeSettings());

            if (installedLanguages.Count == 0)
            {
                menu.Items.Add("No input languages available").Enabled = false;
            }
        }

        private static void OpenWindowsImeSettings()
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "ms-settings:regionlanguage",
                    UseShellExecute = true
                });
            }
            catch
            {
                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = "control.exe",
                        Arguments = "/name Microsoft.Language",
                        UseShellExecute = true
                    });
                }
                catch
                {
                }
            }
        }

        private void AddKnownFolderMenuItem(ToolStripItemCollection items, string label, Environment.SpecialFolder folder, Image image)
        {
            var path = Environment.GetFolderPath(folder);
            if (Directory.Exists(path))
            {
                items.Add(CreateMenuItem(label, image, (_, _) => OpenFolder(path), new ProgramSearchEntry(label, path, IsShellApp: false)));
            }
            else
            {
                image.Dispose();
            }
        }

        private void AddKnownFolderMenuItem(ToolStripItemCollection items, string label, string userProfileChildFolder, Image image)
        {
            var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), userProfileChildFolder);
            if (Directory.Exists(path))
            {
                items.Add(CreateMenuItem(label, image, (_, _) => OpenFolder(path), new ProgramSearchEntry(label, path, IsShellApp: false)));
            }
            else
            {
                image.Dispose();
            }
        }

        private ContextMenuStrip BuildQuickLaunchItemContextMenu()
        {
            var menu = new ContextMenuStrip
            {
                ShowImageMargin = false,
                Font = normalFont
            };
            ApplyWin9xContextMenuStyle(menu);

            var openItem = menu.Items.Add("Open", null, (_, _) => OpenQuickLaunchContextItem());
            var renameItem = menu.Items.Add("Rename", null, (_, _) => RenameQuickLaunchContextItem());
            var moveLeftItem = menu.Items.Add("Move left", null, (_, _) => MoveQuickLaunchContextItem(-1));
            var moveRightItem = menu.Items.Add("Move right", null, (_, _) => MoveQuickLaunchContextItem(1));
            var openContainingFolderItem = menu.Items.Add("Open containing folder", null, (_, _) => OpenQuickLaunchContextItemLocation());
            menu.Items.Add("Open Quick Launch Folder", null, (_, _) => OpenQuickLaunchFolderInExplorerForm());
            menu.Items.Add(new ToolStripSeparator());
            var addSeparatorItem = menu.Items.Add("Add separator", null, (_, _) => AddQuickLaunchSeparator());
            var removeItem = menu.Items.Add("Remove from Quick Launch", null, (_, _) => RemoveQuickLaunchContextItem());
            var propertiesItem = menu.Items.Add("Properties", null, (_, _) => ShowQuickLaunchContextItemProperties());

            menu.Opening += (_, _) =>
            {
                var contextPath = quickLaunchContextShortcutPath;
                var hasFile = !string.IsNullOrWhiteSpace(contextPath) && File.Exists(contextPath);
                var isSeparator = hasFile && Path.GetExtension(contextPath!).Equals(".separator", StringComparison.OrdinalIgnoreCase);
                openItem.Enabled = hasFile && !isSeparator;
                renameItem.Enabled = hasFile;
                moveLeftItem.Enabled = hasFile;
                moveRightItem.Enabled = hasFile;
                openContainingFolderItem.Enabled = hasFile;
                removeItem.Enabled = hasFile;
                propertiesItem.Enabled = hasFile && !isSeparator;
                addSeparatorItem.Enabled = true;
            };

            return menu;
        }

        private void ApplyWin9xContextMenuStyle(ContextMenuStrip menu)
        {
            if (taskbarButtonStyle == TaskbarButtonStyle.Modern)
            {
                // Use default .NET renderer for Modern style
                menu.RenderMode = ToolStripRenderMode.System;
                menu.Renderer = null;
                menu.DropShadowEnabled = true;

                // Remove any custom renderer setup
                foreach (ToolStripItem item in menu.Items)
                {
                    if (item is ToolStripMenuItem menuItem)
                    {
                        RemoveCustomMenuItemRenderer(menuItem);
                    }
                }
            }
            else
            {
                // Use custom Win9x renderer for Win98 style
                menu.RenderMode = ToolStripRenderMode.Professional;
                menu.Renderer = win9xMenuRenderer;
                menu.ShowImageMargin = menu.ShowImageMargin;
                menu.DropShadowEnabled = false;

                // Apply renderer to all submenus when they open
                foreach (ToolStripItem item in menu.Items)
                {
                    if (item is ToolStripMenuItem menuItem)
                    {
                        SetupMenuItemRenderer(menuItem);
                    }
                }
            }
        }

        private void SetupMenuItemRenderer(ToolStripMenuItem menuItem)
        {
            if (menuDropDownOpeningHandlers.ContainsKey(menuItem))
            {
                return;
            }

            // Set up renderer for the dropdown BEFORE it opens
            EventHandler openingHandler = (_, _) =>
            {
                if (menuItem.DropDown is ToolStripDropDown dropDown)
                {
                    dropDown.RenderMode = ToolStripRenderMode.Professional;
                    dropDown.Renderer = win9xMenuRenderer;
                    dropDown.DropShadowEnabled = false;

                    // Recursively apply to all nested menu items
                    foreach (ToolStripItem subItem in dropDown.Items.OfType<ToolStripMenuItem>())
                    {
                        SetupMenuItemRenderer((ToolStripMenuItem)subItem);
                    }
                }
            };
            menuDropDownOpeningHandlers[menuItem] = openingHandler;
            menuItem.DropDownOpening += openingHandler;
        }

        private void RemoveCustomMenuItemRenderer(ToolStripMenuItem menuItem)
        {
            // Remove custom renderer setup - use system default
            if (menuDropDownOpeningHandlers.TryGetValue(menuItem, out var openingHandler))
            {
                menuItem.DropDownOpening -= openingHandler;
                menuDropDownOpeningHandlers.Remove(menuItem);
            }

            foreach (ToolStripItem subItem in menuItem.DropDownItems.OfType<ToolStripMenuItem>())
            {
                RemoveCustomMenuItemRenderer((ToolStripMenuItem)subItem);
            }
        }

        private void SaveTaskbarSettings()
        {
            TaskbarSettingsService.Save(new TaskbarRuntimeSettings(
                StartMenuIconSize: startMenuIconSize,
                TaskIconSize: taskIconSize,
                LazyLoadProgramsSubmenu: lazyLoadProgramsSubmenu,
                PlayVolumeFeedbackSound: playVolumeFeedbackSound,
                UseClassicVolumePopup: useClassicVolumePopup,
                StartMenuSubmenuOpenDelayMs: startMenuSubmenuOpenDelayMs,
                StartMenuRecentProgramsMaxCount: startMenuRecentProgramsMaxCount,
                AutoHideTaskbar: autoHideTaskbar,
                StartOnWindowsStartup: startOnWindowsStartup,
                TaskbarButtonStyle: taskbarButtonStyle.ToString(),
                TaskbarFontName: taskbarFontName,
                TaskbarFontSize: taskbarFontSize,
                TaskbarFontColor: taskbarFontColor,
                TaskbarLocked: taskbarLocked,
                TaskbarRows: taskbarRows,
                TaskbarBaseColor: taskbarBaseColor,
                TaskbarLightColor: taskbarLightColor,
                TaskbarDarkColor: taskbarDarkColor,
                TaskbarBevelSize: taskbarBevelSize,
                ThemeProfileName: selectedThemeProfileName,
                QuickLaunchOrder: GetQuickLaunchOrderFromUi(),
                ShowQuickLaunchToolbar: showQuickLaunchToolbar,
                ShowAddressToolbar: showAddressToolbar,
                AddressToolbarBeforeQuickLaunch: addressToolbarBeforeQuickLaunch,
                ShowSpotifyToolbar: showSpotifyToolbar,
                ShowVolumeToolbar: showVolumeToolbar,
                AddressHistory: addressToolbarHistory.ToList(),
                LastAddressText: addressToolbarComboBox.Text.Trim(),
                StartMenuRecentPrograms: startMenuRecentPrograms
                    .Take(startMenuRecentProgramsMaxCount)
                    .Select(entry => new StartMenuRecentProgramSetting
                    {
                        DisplayName = entry.DisplayName,
                        LaunchPath = entry.LaunchPath,
                        IsShellApp = entry.IsShellApp
                    })
                    .ToList()));
        }

        private void ApplyWindowsStartupSetting(bool enabled)
        {
            WindowsStartupRegistrationService.Apply(enabled, Application.ExecutablePath);
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

        private ToolStripRenderer CreateWin9xMenuRenderer()
        {
            var selectionColor = ResolveMenuSelectionColorForCurrentTheme();
            return new Win9xMenuRendererImpl(taskbarBaseColor, taskbarLightColor, taskbarDarkColor, selectionColor);
        }

        private Color ResolveMenuSelectionColorForCurrentTheme()
        {
            if (TryGetThemeProfile(selectedThemeProfileName, out var profile))
            {
                return profile.MenuSelectionColor;
            }

            return Color.FromArgb(0, 0, 128);
        }

        private bool TryGetThemeProfile(string? name, out ThemeController.ThemeProfile profile)
        {
            return themeController.TryGetProfile(name, out profile);
        }

        private void ApplyThemeProfile(string? profileName, bool preserveCustomizations)
        {
            if (!TryGetThemeProfile(profileName, out var profile))
            {
                selectedThemeProfileName = "Custom";
                return;
            }

            selectedThemeProfileName = profile.Name;
            if (preserveCustomizations)
            {
                return;
            }

            if (Enum.TryParse<TaskbarButtonStyle>(profile.ButtonStyle, ignoreCase: true, out var parsedStyle))
            {
                taskbarButtonStyle = parsedStyle;
            }
            taskbarBevelSize = profile.BevelSize;
            taskbarFontName = profile.FontName;
            taskbarFontSize = profile.FontSize;
            taskbarFontColor = profile.FontColor;
            taskbarBaseColor = profile.BaseColor;
            taskbarLightColor = profile.LightColor;
            taskbarDarkColor = profile.DarkColor;
        }

        private void RefreshTaskbarFonts()
        {
            normalFont.Dispose();
            activeFont.Dispose();
            normalFont = CreateTaskbarFont(FontStyle.Regular);
            activeFont = CreateTaskbarFont(FontStyle.Bold);

            Font = normalFont;
            lblClock.Font = normalFont;
            lblClock.ForeColor = taskbarFontColor;
            UpdateClockDisplay();
            languageIndicatorLabel.Font = InputIndicatorFont;
            languageIndicatorLabel.ForeColor = Color.White;
            languageIndicatorLabel.Size = new Size(18, 18);
            languageIndicatorLabel.MinimumSize = new Size(18, 18);
            languageIndicatorLabel.MaximumSize = new Size(18, 18);
            imeIndicatorLabel.Font = InputIndicatorFont;
            imeIndicatorLabel.ForeColor = Color.White;
            imeIndicatorLabel.Size = new Size(18, 18);
            imeIndicatorLabel.MinimumSize = new Size(18, 18);
            imeIndicatorLabel.MaximumSize = new Size(18, 18);
            btnStart.Font = activeFont;
            btnStart.ForeColor = taskbarFontColor;
            btnStart.Height = TaskbarButtonHeight;

            taskWindowMenu.Font = normalFont;
            taskbarContextMenu.Font = normalFont;
            clockContextMenu.Font = normalFont;
            languageIndicatorContextMenu.Font = normalFont;
            imeIndicatorContextMenu.Font = normalFont;
            quickLaunchItemContextMenu.Font = normalFont;
            startMenu.Font = normalFont;
            addressToolbarComboBox.Font = normalFont;
            addressToolbarComboBox.Height = Math.Max(18, TextRenderer.MeasureText("My Computer", normalFont).Height + 8);
            addressInputHostPanel.Height = GetAddressInputHostHeight();
            CenterAddressTextBoxVertically();

            // Update all task buttons
            foreach (var button in taskButtons.Values)
            {
                button.Font = normalFont;
                button.ForeColor = taskbarFontColor;
            }

            // Update all Quick Launch buttons
            foreach (Control control in quickLaunchPanel.Controls)
            {
                if (control is Button button)
                {
                    button.Font = normalFont;
                    button.ForeColor = taskbarFontColor;
                }
            }

            btnSpotifyPrevious.Font = normalFont;
            btnSpotifyPrevious.ForeColor = taskbarFontColor;
            btnSpotifyPlayPause.Font = normalFont;
            btnSpotifyPlayPause.ForeColor = taskbarFontColor;
            btnSpotifyNext.Font = normalFont;
            btnSpotifyNext.ForeColor = taskbarFontColor;
            spotifyNowPlayingLabel.Font = normalFont;
            spotifyNowPlayingLabel.ForeColor = taskbarFontColor;
        }

        private void RefreshTaskbarColors()
        {
            BackColor = taskbarBaseColor;
            taskbarResizeHandlePanel.BackColor = taskbarBaseColor;
            contentPanel.BackColor = taskbarBaseColor;
            startHostPanel.BackColor = taskbarBaseColor;
            taskButtonsHostPanel.BackColor = taskbarBaseColor;
            addressRowHostPanel.BackColor = taskbarBaseColor;
            taskButtonsPanel.BackColor = taskbarBaseColor;
            quickLaunchPanel.BackColor = taskbarBaseColor;
            addressToolbarPanel.BackColor = taskbarBaseColor;
            volumeToolbarPanel.BackColor = taskbarBaseColor;
            spotifyToolbarPanel.BackColor = taskbarBaseColor;
            spotifyControlsPanel.BackColor = taskbarBaseColor;
            spotifyNowPlayingLabel.BackColor = taskbarBaseColor;
            trayIconsPanel.BackColor = taskbarBaseColor;
            notificationAreaPanel.BackColor = taskbarBaseColor;
            lblClock.BackColor = taskbarBaseColor;
            customVolumeIcon.BackColor = taskbarBaseColor;
            toolbarsSeparatorPanel.BackColor = taskbarBaseColor;
            spotifyToolbarSeparatorPanel.BackColor = taskbarBaseColor;
            taskButtonsSeparatorPanel.BackColor = taskbarBaseColor;
            addressRowSeparatorPanel.BackColor = taskbarBaseColor;
            startToolbarSeparatorPanel.BackColor = taskbarBaseColor;
            GripPanel.ApplyTheme(taskbarBaseColor, quickLaunchGripPanel, addressToolbarGripPanel, addressRowGripPanel, volumeToolbarGripPanel, spotifyToolbarGripPanel, runningProgramsGripPanel);
            addressToolbarComboBox.ForeColor = taskbarFontColor;
            addressToolbarComboBox.BackColor = Color.White;
            ApplyClassicAddressComboStyle();
            languageIndicatorLabel.BackColor = InputIndicatorBackColor;
            imeIndicatorLabel.BackColor = InputIndicatorBackColor;

            win9xMenuRenderer = CreateWin9xMenuRenderer();
            ApplyWin9xContextMenuStyle(startMenu);
            ApplyWin9xContextMenuStyle(taskWindowMenu);
            ApplyWin9xContextMenuStyle(taskbarContextMenu);
            ApplyWin9xContextMenuStyle(clockContextMenu);
            ApplyWin9xContextMenuStyle(languageIndicatorContextMenu);
            ApplyWin9xContextMenuStyle(imeIndicatorContextMenu);
            ApplyWin9xContextMenuStyle(quickLaunchItemContextMenu);

            UpdateResizeUiState();
            notificationAreaPanel.Invalidate();
            taskbarResizeHandlePanel.Invalidate();
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
                    button.Padding = new Padding(0);
                    break;
                case TaskbarButtonStyle.Modern:
                default:
                    button.FlatStyle = FlatStyle.Standard;
                    button.UseVisualStyleBackColor = true;
                    button.Padding = new Padding(1); // +1 padding for Modern
                    button.Margin = new Padding(0, 0, TaskbarButtonGap, 0);
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

            var rect = new Rectangle(0, 0, button.Width - 1, button.Height - 1);
            if (rect.Width < 3 || rect.Height < 3)
            {
                return;
            }

            var isMousePressed = button.Capture && button.ClientRectangle.Contains(button.PointToClient(Cursor.Position));
            var isSelectedPressed = IsButtonSelectedPressed(button);
            var isPressed = isMousePressed || isSelectedPressed;
            var innerDark = taskbarDarkColor;
            var outerDark = Color.Black;

            if (isPressed)
            {
                // Pressed: invert the bevel to appear sunken.
                using var topLeftInnerPen = new Pen(innerDark);
                using var topLeftOuterPen = new Pen(outerDark);
                using var bottomRightPen = new Pen(taskbarLightColor);
                e.Graphics.DrawLine(topLeftOuterPen, 0, 0, rect.Right, 0);
                e.Graphics.DrawLine(topLeftOuterPen, 0, 0, 0, rect.Bottom);
                e.Graphics.DrawLine(topLeftInnerPen, 1, 1, rect.Right - 1, 1);
                e.Graphics.DrawLine(topLeftInnerPen, 1, 1, 1, rect.Bottom - 1);
                e.Graphics.DrawLine(bottomRightPen, rect.Right, 1, rect.Right, rect.Bottom);
                e.Graphics.DrawLine(bottomRightPen, 1, rect.Bottom, rect.Right, rect.Bottom);
            }
            else
            {
                // Classic Win98 bevel:
                // top/left: 1px light; right/bottom: inner 1px dark + outer 1px black.
                using var lightPen = new Pen(taskbarLightColor);
                using var innerDarkPen = new Pen(innerDark);
                using var outerDarkPen = new Pen(outerDark);

                e.Graphics.DrawLine(lightPen, 0, 0, rect.Right - 2, 0);
                e.Graphics.DrawLine(lightPen, 0, 0, 0, rect.Bottom - 2);
                e.Graphics.DrawLine(innerDarkPen, rect.Right - 1, 0, rect.Right - 1, rect.Bottom - 1);
                e.Graphics.DrawLine(innerDarkPen, 0, rect.Bottom - 1, rect.Right - 1, rect.Bottom - 1);
                e.Graphics.DrawLine(outerDarkPen, rect.Right, 0, rect.Right, rect.Bottom);
                e.Graphics.DrawLine(outerDarkPen, 0, rect.Bottom, rect.Right, rect.Bottom);
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
                return !window.IsMinimized && handle == interactionStateMachine.ActiveWindowHandle;
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

            quickLaunchDragSourcePath = null;
            ClearQuickLaunchDragPreview();

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
                var settings = TaskbarSettingsStore.Load();
                var entries = quickLaunchController.GetOrderedEntries(quickLaunchFolder, settings.QuickLaunchOrder);
                foreach (var entry in entries)
                {
                    if (entry.IsSeparator)
                    {
                        quickLaunchPanel.Controls.Add(CreateQuickLaunchSeparator(entry.FullPath));
                    }
                    else
                    {
                        quickLaunchPanel.Controls.Add(CreateQuickLaunchButton(entry.FullPath));
                    }
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
            button.MouseDown += QuickLaunchButton_MouseDown;
            button.MouseMove += QuickLaunchButton_MouseMove;
            return button;
        }

        private Control CreateQuickLaunchSeparator(string separatorPath)
        {
            var separator = new Panel
            {
                Width = 8,
                Height = QuickLaunchButtonSize,
                Margin = new Padding(0, 0, TaskbarButtonGap, 0),
                Tag = separatorPath,
                Cursor = Cursors.SizeWE
            };

            separator.Paint += (_, e) =>
            {
                var x = separator.Width / 2;
                using var darkPen = new Pen(taskbarDarkColor);
                using var lightPen = new Pen(taskbarLightColor);
                e.Graphics.DrawLine(darkPen, x - 1, 2, x - 1, separator.Height - 3);
                e.Graphics.DrawLine(lightPen, x, 2, x, separator.Height - 3);
            };

            separator.MouseUp += QuickLaunchSeparator_MouseUp;
            separator.MouseDown += QuickLaunchSeparator_MouseDown;
            separator.MouseMove += QuickLaunchSeparator_MouseMove;
            quickLaunchToolTip.SetToolTip(separator, "Quick Launch separator");
            return separator;
        }

        private Image? GetQuickLaunchButtonIcon(string quickLaunchPath)
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

            return GetSmallShellIcon(quickLaunchPath, isDirectory: false)
                ?? LoadMenuAsset("Application Window.png", SystemIcons.Application.ToBitmap());
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

        private int GetQuickLaunchRows()
        {
            return Math.Clamp(taskbarRows, 1, 3);
        }

        private void ResizeQuickLaunchPanel()
        {
            var controls = quickLaunchPanel.Controls.OfType<Control>().ToList();
            if (controls.Count == 0)
            {
                quickLaunchPanel.Width = 0;
                return;
            }

            var rows = GetQuickLaunchRows();
            quickLaunchPanel.FlowDirection = rows > 1 ? FlowDirection.TopDown : FlowDirection.LeftToRight;
            quickLaunchPanel.WrapContents = rows > 1;

            foreach (var control in controls)
            {
                control.Margin = rows > 1
                    ? new Padding(0, 0, TaskbarButtonGap, TaskButtonRowBottomGap)
                    : new Padding(0, 0, TaskbarButtonGap, 0);
            }

            var autoWidth = rows <= 1
                ? controls.Sum(control => control.Width + control.Margin.Horizontal) + quickLaunchPanel.Padding.Horizontal
                : (Math.Max(1, (int)Math.Ceiling(controls.Count / (double)rows)) * (QuickLaunchButtonSize + TaskbarButtonGap)) + quickLaunchPanel.Padding.Horizontal;

            quickLaunchPanel.Width = autoWidth;
            quickLaunchPanel.PerformLayout();
        }

        private List<string> GetQuickLaunchOrderFromUi()
        {
            return quickLaunchController.GetOrderFromPanel(quickLaunchPanel);
        }

        private void QuickLaunchButton_Click(object? sender, EventArgs e)
        {
            if (sender is not Button button || button.Tag is not string shortcutPath)
            {
                return;
            }

            if (string.Equals(Path.GetExtension(shortcutPath), ".separator", StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            LaunchQuickLaunchTarget(shortcutPath);
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

            switch (taskbarButtonStyle)
            {
                case TaskbarButtonStyle.Win98:
                    button.FlatStyle = FlatStyle.Flat;
                    button.FlatAppearance.BorderSize = 0;
                    button.FlatAppearance.MouseOverBackColor = taskbarBaseColor;
                    button.FlatAppearance.MouseDownBackColor = taskbarBaseColor;
                    button.UseVisualStyleBackColor = false;
                    button.BackColor = taskbarBaseColor;
                    button.Padding = new Padding(0);
                    button.Paint += QuickLaunchButton_Paint;
                    button.MouseDown += QuickLaunchButton_MouseDownVisual;
                    button.MouseUp += QuickLaunchButton_MouseUpVisual;
                    button.MouseLeave += QuickLaunchButton_StateChangedVisual;
                    button.LostFocus += QuickLaunchButton_StateChangedVisual;
                    button.MouseCaptureChanged += QuickLaunchButton_StateChangedVisual;
                    break;
                case TaskbarButtonStyle.Modern:
                default:
                    button.FlatStyle = FlatStyle.Standard;
                    button.UseVisualStyleBackColor = true;
                    button.Padding = new Padding(1); // +1 padding for Modern
                    break;
            }
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

            var isHovered = button.ClientRectangle.Contains(button.PointToClient(Cursor.Position));
            var isPressed = button.Capture && isHovered;
            var rect = new Rectangle(0, 0, button.Width - 1, button.Height - 1);

            if (isPressed)
            {
                ControlPaint.DrawBorder(e.Graphics, rect,
                    taskbarDarkColor, 1, ButtonBorderStyle.Solid,
                    taskbarDarkColor, 1, ButtonBorderStyle.Solid,
                    taskbarLightColor, 1, ButtonBorderStyle.Solid,
                    taskbarLightColor, 1, ButtonBorderStyle.Solid);
            }
            else if (isHovered)
            {
                // Hover: 1px raised bevel (??.
                ControlPaint.DrawBorder(e.Graphics, rect,
                    taskbarLightColor, 1, ButtonBorderStyle.Solid,
                    taskbarLightColor, 1, ButtonBorderStyle.Solid,
                    taskbarDarkColor, 1, ButtonBorderStyle.Solid,
                    taskbarDarkColor, 1, ButtonBorderStyle.Solid);
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
            if (e.Button == MouseButtons.Left)
            {
                quickLaunchDragSourcePath = null;
                return;
            }

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

            LaunchQuickLaunchTarget(quickLaunchContextShortcutPath);
        }

        private void LaunchQuickLaunchTarget(string shortcutPath)
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = shortcutPath,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to launch quick launch target '{shortcutPath}': {ex.Message}");
                MessageBox.Show(
                    this,
                    $"Unable to launch '{Path.GetFileNameWithoutExtension(shortcutPath)}'.",
                    "Quick Launch",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
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
                SaveTaskbarSettings();
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

        private void OpenQuickLaunchContextItemLocation()
        {
            if (string.IsNullOrWhiteSpace(quickLaunchContextShortcutPath) || !File.Exists(quickLaunchContextShortcutPath))
            {
                return;
            }

            LaunchProcess("explorer.exe", $"/select,\"{quickLaunchContextShortcutPath}\"");
        }

        private void RenameQuickLaunchContextItem()
        {
            if (string.IsNullOrWhiteSpace(quickLaunchContextShortcutPath) || !File.Exists(quickLaunchContextShortcutPath))
            {
                return;
            }

            try
            {
                var originalPath = quickLaunchContextShortcutPath;
                var originalName = Path.GetFileNameWithoutExtension(originalPath);
                var extension = Path.GetExtension(originalPath);
                var input = Microsoft.VisualBasic.Interaction.InputBox("Enter a new name:", "Rename Quick Launch Item", originalName);
                if (string.IsNullOrWhiteSpace(input) || input.Equals(originalName, StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }

                var safeName = string.Join("_", input.Split(Path.GetInvalidFileNameChars(), StringSplitOptions.RemoveEmptyEntries)).Trim();
                if (string.IsNullOrWhiteSpace(safeName))
                {
                    return;
                }

                var destination = Path.Combine(Path.GetDirectoryName(originalPath) ?? string.Empty, safeName + extension);
                destination = GetUniqueFilePath(Path.GetDirectoryName(destination) ?? string.Empty, Path.GetFileName(destination));
                File.Move(originalPath, destination);
                quickLaunchContextShortcutPath = destination;
                LoadQuickLaunchButtons();
                SaveTaskbarSettings();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to rename quick launch item: {ex.Message}");
            }
        }

        private void MoveQuickLaunchContextItem(int direction)
        {
            if (string.IsNullOrWhiteSpace(quickLaunchContextShortcutPath) || !File.Exists(quickLaunchContextShortcutPath))
            {
                return;
            }

            var controls = quickLaunchPanel.Controls.Cast<Control>().ToList();
            var currentIndex = controls.FindIndex(control => string.Equals(control.Tag as string, quickLaunchContextShortcutPath, StringComparison.OrdinalIgnoreCase));
            if (currentIndex < 0)
            {
                return;
            }

            var targetIndex = Math.Clamp(currentIndex + direction, 0, controls.Count - 1);
            if (targetIndex == currentIndex)
            {
                return;
            }

            var control = controls[currentIndex];
            quickLaunchPanel.Controls.SetChildIndex(control, targetIndex);
            quickLaunchPanel.Invalidate();
            ResizeQuickLaunchPanel();
            SaveTaskbarSettings();
        }

        private void AddQuickLaunchSeparator()
        {
            try
            {
                var quickLaunchFolder = GetQuickLaunchFolderPath();
                Directory.CreateDirectory(quickLaunchFolder);
                var separatorPath = GetUniqueFilePath(quickLaunchFolder, "Separator.separator");
                File.WriteAllText(separatorPath, string.Empty);
                LoadQuickLaunchButtons();
                SaveTaskbarSettings();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to add quick launch separator: {ex.Message}");
            }
        }

        private void QuickLaunchButton_MouseDown(object? sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left || sender is not Button button || button.Tag is not string path)
            {
                return;
            }

            quickLaunchDragSourcePath = path;
            quickLaunchDragStartPoint = e.Location;
        }

        private void QuickLaunchButton_MouseMove(object? sender, MouseEventArgs e)
        {
            if ((Control.MouseButtons & MouseButtons.Left) == 0 || string.IsNullOrWhiteSpace(quickLaunchDragSourcePath) || sender is not Control control)
            {
                return;
            }

            if (Math.Abs(e.X - quickLaunchDragStartPoint.X) < SystemInformation.DragSize.Width / 2 &&
                Math.Abs(e.Y - quickLaunchDragStartPoint.Y) < SystemInformation.DragSize.Height / 2)
            {
                return;
            }

            var data = new DataObject();
            data.SetData("QuickLaunchItemPath", quickLaunchDragSourcePath);
            control.DoDragDrop(data, DragDropEffects.Move);
        }

        private void QuickLaunchSeparator_MouseDown(object? sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left || sender is not Control control || control.Tag is not string path)
            {
                return;
            }

            quickLaunchDragSourcePath = path;
            quickLaunchDragStartPoint = e.Location;
        }

        private void QuickLaunchSeparator_MouseMove(object? sender, MouseEventArgs e)
        {
            if ((Control.MouseButtons & MouseButtons.Left) == 0 || string.IsNullOrWhiteSpace(quickLaunchDragSourcePath) || sender is not Control control)
            {
                return;
            }

            if (Math.Abs(e.X - quickLaunchDragStartPoint.X) < SystemInformation.DragSize.Width / 2 &&
                Math.Abs(e.Y - quickLaunchDragStartPoint.Y) < SystemInformation.DragSize.Height / 2)
            {
                return;
            }

            var data = new DataObject();
            data.SetData("QuickLaunchItemPath", quickLaunchDragSourcePath);
            control.DoDragDrop(data, DragDropEffects.Move);
        }

        private void QuickLaunchSeparator_MouseUp(object? sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                quickLaunchDragSourcePath = null;
                return;
            }

            if (e.Button != MouseButtons.Right || sender is not Control control || control.Tag is not string shortcutPath)
            {
                return;
            }

            quickLaunchContextShortcutPath = shortcutPath;
            quickLaunchItemContextMenu.Show(control, e.Location);
        }

        private void ReorderQuickLaunchItemByDropPoint(string sourcePath, Point dropPoint)
        {
            var controls = quickLaunchPanel.Controls.Cast<Control>().ToList();
            var sourceControl = controls.FirstOrDefault(control => string.Equals(control.Tag as string, sourcePath, StringComparison.OrdinalIgnoreCase));
            if (sourceControl == null)
            {
                return;
            }

            var targetIndex = GetQuickLaunchDropInsertPosition(dropPoint);

            var sourceIndex = quickLaunchPanel.Controls.GetChildIndex(sourceControl);
            if (sourceIndex < targetIndex)
            {
                targetIndex--;
            }

            targetIndex = Math.Clamp(targetIndex, 0, quickLaunchPanel.Controls.Count - 1);
            quickLaunchPanel.Controls.SetChildIndex(sourceControl, targetIndex);
            quickLaunchPanel.Invalidate();
            ResizeQuickLaunchPanel();
            SaveTaskbarSettings();
        }

        private void QuickLaunchPanel_DragEnter(object? sender, DragEventArgs e)
        {
            if (e.Data?.GetDataPresent("QuickLaunchItemPath") == true)
            {
                e.Effect = DragDropEffects.Move;
                return;
            }

            e.Effect = e.Data?.GetDataPresent(DataFormats.FileDrop) == true
                ? DragDropEffects.Copy
                : DragDropEffects.None;
        }

        private void QuickLaunchPanel_DragOver(object? sender, DragEventArgs e)
        {
            if (e.Data?.GetDataPresent("QuickLaunchItemPath") != true)
            {
                ClearQuickLaunchDragPreview();
                return;
            }

            e.Effect = DragDropEffects.Move;
            var sourcePath = e.Data.GetData("QuickLaunchItemPath") as string;
            var dropPoint = quickLaunchPanel.PointToClient(new Point(e.X, e.Y));
            UpdateQuickLaunchDragPreview(dropPoint, sourcePath);
        }

        private void QuickLaunchPanel_DragLeave(object? sender, EventArgs e)
        {
            ClearQuickLaunchDragPreview();
        }

        private void QuickLaunchPanel_Paint(object? sender, PaintEventArgs e)
        {
            if (quickLaunchPreviewInsertPosition < 0)
            {
                return;
            }

            var x = GetQuickLaunchPreviewX(quickLaunchPreviewInsertPosition);
            x = Math.Clamp(x, 0, Math.Max(0, quickLaunchPanel.ClientSize.Width - 1));
            var top = 2;
            var bottom = Math.Max(top, quickLaunchPanel.ClientSize.Height - 3);

            using var lightPen = new Pen(taskbarLightColor);
            using var darkPen = new Pen(taskbarDarkColor);
            e.Graphics.DrawLine(lightPen, x, top, x, bottom);
            if (x + 1 < quickLaunchPanel.ClientSize.Width)
            {
                e.Graphics.DrawLine(darkPen, x + 1, top, x + 1, bottom);
            }
        }

        private void UpdateQuickLaunchDragPreview(Point dropPoint, string? sourcePath)
        {
            var insertPosition = GetQuickLaunchDropInsertPosition(dropPoint);
            if (!string.IsNullOrWhiteSpace(sourcePath))
            {
                var sourceControl = quickLaunchPanel.Controls
                    .OfType<Control>()
                    .FirstOrDefault(control => string.Equals(control.Tag as string, sourcePath, StringComparison.OrdinalIgnoreCase));

                if (sourceControl != null)
                {
                    var sourceIndex = quickLaunchPanel.Controls.GetChildIndex(sourceControl);
                    if (insertPosition > sourceIndex)
                    {
                        insertPosition--;
                    }
                }
            }

            insertPosition = Math.Clamp(insertPosition, 0, quickLaunchPanel.Controls.Count);
            if (quickLaunchPreviewInsertPosition == insertPosition)
            {
                return;
            }

            quickLaunchPreviewInsertPosition = insertPosition;
            quickLaunchPanel.Invalidate();
        }

        private void ClearQuickLaunchDragPreview()
        {
            if (quickLaunchPreviewInsertPosition < 0)
            {
                return;
            }

            quickLaunchPreviewInsertPosition = -1;
            quickLaunchPanel.Invalidate();
        }

        private int GetQuickLaunchDropInsertPosition(Point dropPoint)
        {
            var controls = quickLaunchPanel.Controls.Cast<Control>().ToList();
            return quickLaunchController.GetDropInsertPosition(controls, dropPoint);
        }

        private int GetQuickLaunchPreviewX(int insertPosition)
        {
            var controls = quickLaunchPanel.Controls.Cast<Control>().ToList();
            if (controls.Count == 0 || insertPosition <= 0)
            {
                return 0;
            }

            if (insertPosition >= controls.Count)
            {
                var last = controls[^1];
                return last.Right + Math.Max(1, last.Margin.Right / 2);
            }

            var next = controls[insertPosition];
            return Math.Max(0, next.Left - Math.Max(1, next.Margin.Left / 2));
        }

        private void QuickLaunchPanel_DragDrop(object? sender, DragEventArgs e)
        {
            if (e.Data?.GetDataPresent("QuickLaunchItemPath") == true)
            {
                var sourcePath = e.Data.GetData("QuickLaunchItemPath") as string;
                if (!string.IsNullOrWhiteSpace(sourcePath))
                {
                    ReorderQuickLaunchItemByDropPoint(sourcePath, quickLaunchPanel.PointToClient(new Point(e.X, e.Y)));
                }

                ClearQuickLaunchDragPreview();
                quickLaunchDragSourcePath = null;

                return;
            }

            ClearQuickLaunchDragPreview();

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
                SaveTaskbarSettings();
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
            if (windowsMinimizedByShowDesktop.Count > 0)
            {
                RestoreShowDesktopWindows();
                return;
            }

            windowsMinimizedByShowDesktop.Clear();
            foreach (var window in windowsByHandle.Values)
            {
                if (!window.IsMinimized)
                {
                    windowsMinimizedByShowDesktop.Add(window.Handle);
                    window.Minimize();
                }
            }
        }

        private void RestoreShowDesktopWindows()
        {
            foreach (var handle in windowsMinimizedByShowDesktop.ToList())
            {
                if (windowsByHandle.TryGetValue(handle, out var window) && window.IsMinimized)
                {
                    window.Restore();
                }
            }

            windowsMinimizedByShowDesktop.Clear();
            RefreshTaskbar();
        }

        private void RestoreMinimizedWindows()
        {
            foreach (var window in windowsByHandle.Values)
            {
                if (window.IsMinimized)
                {
                    window.Restore();
                }
            }

            windowsMinimizedByShowDesktop.Clear();
            RefreshTaskbar();
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

        private void InputIndicatorLabel_MouseUp(object? sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left && e.Button != MouseButtons.Right)
            {
                return;
            }

            if (sender == languageIndicatorLabel)
            {
                languageIndicatorContextMenu.Show(languageIndicatorLabel, new Point(0, languageIndicatorLabel.Height));
                return;
            }

            if (sender == imeIndicatorLabel)
            {
                imeIndicatorContextMenu.Show(imeIndicatorLabel, new Point(0, imeIndicatorLabel.Height));
            }
        }

        private void ShowTimeDetailsModal()
        {
            var popup = EnsureTimeDetailsPopup();
            var clockBounds = lblClock.RectangleToScreen(lblClock.ClientRectangle);
            var popupX = clockBounds.Right - popup.Width;
            var popupY = clockBounds.Top - popup.Height - 4;
            var screenBounds = Screen.FromPoint(clockBounds.Location).WorkingArea;

            if (popupX < screenBounds.Left)
            {
                popupX = screenBounds.Left;
            }

            if (popupX + popup.Width > screenBounds.Right)
            {
                popupX = screenBounds.Right - popup.Width;
            }

            if (popupY < screenBounds.Top)
            {
                popupY = clockBounds.Bottom + 4;
            }

            if (popup.Visible)
            {
                popup.Hide();
                return;
            }

            popup.ShowAt(new Point(popupX, popupY));
        }

        private TimeDetailsForm EnsureTimeDetailsPopup()
        {
            if (timeDetailsPopup is null || timeDetailsPopup.IsDisposed)
            {
                timeDetailsPopup = new TimeDetailsForm();
            }

            return timeDetailsPopup;
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

        private static void OpenEnvironmentVariablesSettings()
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "rundll32.exe",
                    Arguments = "sysdm.cpl,EditEnvironmentVariables",
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
                    FileName = "SystemPropertiesAdvanced.exe",
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to open Environment Variables settings: {ex.Message}");
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
                for (var i = 0; i < lines && i < taskbarResizeHandlePanel.Height; i++)
                {
                    e.Graphics.DrawLine(lightPen, 0, i, taskbarResizeHandlePanel.Width, i);
                }

                return;
            }

            using var darkPen = new Pen(taskbarDarkColor);
            var lightLines = Math.Max(1, lines - 1);
            for (var i = 0; i < lightLines && i < taskbarResizeHandlePanel.Height; i++)
            {
                e.Graphics.DrawLine(lightPen, 0, i, taskbarResizeHandlePanel.Width, i);
            }

            for (var i = lightLines; i < lines; i++)
            {
                if (i >= taskbarResizeHandlePanel.Height)
                {
                    break;
                }

                e.Graphics.DrawLine(darkPen, 0, i, taskbarResizeHandlePanel.Width, i);
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
            taskbarResizeHandlePanel.Visible = true;
            taskbarResizeHandlePanel.Height = ResizeGripHeight;
            taskbarResizeHandlePanel.Cursor = taskbarLocked ? Cursors.Default : Cursors.SizeNS;
            taskbarResizeHandlePanel.Invalidate();
            GripPanel.ApplyLockState(taskbarLocked, quickLaunchGripPanel, addressToolbarGripPanel, addressRowGripPanel, volumeToolbarGripPanel, spotifyToolbarGripPanel, runningProgramsGripPanel);
        }

        private void SetTaskbarLocked(bool locked)
        {
            taskbarLocked = locked;
            UpdateResizeUiState();
            ApplyToolbarLayout(refreshBounds: false);
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

            var resizeHotZone = ResizeGripHeight;
            if (e.Y >= resizeHotZone)
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
                ResizeTrayPanel();
                RefreshTrayHostPlacement();
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
            var originalState = CaptureOptionsState();
            using var dialog = new TaskbarOptionsForm(
                startMenuIconSize,
                taskIconSize,
                lazyLoadProgramsSubmenu,
                playVolumeFeedbackSound,
                useClassicVolumePopup,
                startMenuSubmenuOpenDelayMs,
                startMenuRecentProgramsMaxCount,
                autoHideTaskbar,
                startOnWindowsStartup,
                taskbarButtonStyle.ToString(),
                taskbarBevelSize,
                taskbarFontName,
                taskbarFontSize,
                taskbarFontColor,
                taskbarBaseColor,
                taskbarLightColor,
                taskbarDarkColor,
                selectedThemeProfileName,
                OpenQuickLaunchFolderInExplorerForm);
            dialog.ApplyRequested += () => ApplyOptionsFromDialog(dialog, persist: false);
            if (dialog.ShowDialog(this) != DialogResult.OK)
            {
                RestoreOptionsState(originalState, persist: false);
                return;
            }

            ApplyOptionsFromDialog(dialog, persist: true);
        }

        private void ShowDiagnosticsPanel()
        {
            if (diagnosticsPanel != null && !diagnosticsPanel.IsDisposed)
            {
                diagnosticsPanel.Show();
                diagnosticsPanel.BringToFront();
                diagnosticsPanel.Activate();
                return;
            }

            diagnosticsPanel = new DiagnosticsPanelForm(CreateDiagnosticsSnapshot);
            diagnosticsPanel.FormClosed += (_, _) => diagnosticsPanel = null;
            diagnosticsPanel.Show(this);
        }

        private DiagnosticsPanelForm.DiagnosticsSnapshot CreateDiagnosticsSnapshot()
        {
            var trayIcons = shellManager.NotificationArea.TrayIcons;
            var visibleTrayIcons = trayIcons.Count(icon => !icon.IsHidden && icon.Icon != null && !ShouldHideTrayIcon(icon));

            return new DiagnosticsPanelForm.DiagnosticsSnapshot(
                ActiveHandle: $"0x{interactionStateMachine.ActiveWindowHandle.ToInt64():X}",
                ForegroundHandleBeforeClick: $"0x{interactionStateMachine.ForegroundWindowBeforeTaskClick.ToInt64():X}",
                WindowCount: windowsByHandle.Count,
                TaskButtonCount: taskButtons.Count,
                TrayIconCount: trayIcons.Count,
                VisibleTrayIconCount: visibleTrayIcons,
                ProgramCacheEntries: programFolderCache.Count,
                SearchCacheEntries: programSearchEntriesCache.Count,
                QuickLaunchControlCount: quickLaunchPanel.Controls.Count,
                ThemeProfile: selectedThemeProfileName,
                Timestamp: DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"));
        }

        private OptionsState CaptureOptionsState()
        {
            return new OptionsState(
                startMenuIconSize,
                taskIconSize,
                lazyLoadProgramsSubmenu,
                playVolumeFeedbackSound,
                useClassicVolumePopup,
                startMenuSubmenuOpenDelayMs,
                startMenuRecentProgramsMaxCount,
                autoHideTaskbar,
                startOnWindowsStartup,
                taskbarButtonStyle,
                taskbarBevelSize,
                taskbarFontName,
                taskbarFontSize,
                taskbarFontColor,
                taskbarBaseColor,
                taskbarLightColor,
                taskbarDarkColor,
                selectedThemeProfileName);
        }

        private void RestoreOptionsState(OptionsState state, bool persist)
        {
            startMenuIconSize = state.StartMenuIconSize;
            taskIconSize = state.TaskIconSize;
            lazyLoadProgramsSubmenu = state.LazyLoadProgramsSubmenu;
            playVolumeFeedbackSound = state.PlayVolumeFeedbackSound;
            useClassicVolumePopup = state.UseClassicVolumePopup;
            startMenuSubmenuOpenDelayMs = state.StartMenuSubmenuOpenDelayMs;
            startMenuRecentProgramsMaxCount = state.StartMenuRecentProgramsMaxCount;
            autoHideTaskbar = state.AutoHideTaskbar;
            startOnWindowsStartup = state.StartOnWindowsStartup;
            taskbarButtonStyle = state.ButtonStyle;
            taskbarBevelSize = state.BevelSize;
            taskbarFontName = state.FontName;
            taskbarFontSize = state.FontSize;
            taskbarFontColor = state.FontColor;
            taskbarBaseColor = state.BaseColor;
            taskbarLightColor = state.LightColor;
            taskbarDarkColor = state.DarkColor;
            selectedThemeProfileName = state.ThemeProfileName;
            ApplyWindowsStartupSetting(startOnWindowsStartup);

            volumePopup.PlayFeedbackSoundOnMouseUp = playVolumeFeedbackSound;
            if (startMenuRecentPrograms.Count > startMenuRecentProgramsMaxCount)
            {
                startMenuRecentPrograms.RemoveRange(startMenuRecentProgramsMaxCount, startMenuRecentPrograms.Count - startMenuRecentProgramsMaxCount);
            }
            RefreshStartMenuRecentProgramsSection();
            ApplyIconSizeSettings();

            if (persist)
            {
                SaveTaskbarSettings();
            }
        }

        private void ApplyOptionsFromDialog(TaskbarOptionsForm dialog, bool persist)
        {
            startMenuIconSize = dialog.StartMenuIconSize;
            taskIconSize = dialog.TaskIconSize;
            lazyLoadProgramsSubmenu = dialog.LazyLoadProgramsSubmenu;
            playVolumeFeedbackSound = dialog.PlayVolumeFeedbackSound;
            useClassicVolumePopup = dialog.UseClassicVolumePopup;
            startMenuSubmenuOpenDelayMs = dialog.StartMenuSubmenuOpenDelayMs;
            startMenuRecentProgramsMaxCount = dialog.StartMenuRecentProgramsMaxCount;
            autoHideTaskbar = dialog.AutoHideTaskbar;
            startOnWindowsStartup = dialog.StartOnWindowsStartup;
            taskbarBevelSize = Math.Clamp(dialog.TaskbarBevelSize, 1, 4);
            taskbarFontName = dialog.TaskbarFontName;
            taskbarFontSize = dialog.TaskbarFontSize;
            taskbarFontColor = dialog.TaskbarFontColor;
            taskbarBaseColor = dialog.TaskbarBaseColor;
            taskbarLightColor = dialog.TaskbarLightColor;
            taskbarDarkColor = dialog.TaskbarDarkColor;
            if (Enum.TryParse<TaskbarButtonStyle>(dialog.TaskbarButtonStyle, ignoreCase: true, out var parsedStyle))
            {
                taskbarButtonStyle = parsedStyle;
            }

            selectedThemeProfileName = dialog.SelectedThemeProfileName;
            ApplyThemeProfile(selectedThemeProfileName, preserveCustomizations: selectedThemeProfileName.Equals("Custom", StringComparison.OrdinalIgnoreCase));
            ApplyWindowsStartupSetting(startOnWindowsStartup);
            volumePopup.PlayFeedbackSoundOnMouseUp = playVolumeFeedbackSound;
            if (startMenuRecentPrograms.Count > startMenuRecentProgramsMaxCount)
            {
                startMenuRecentPrograms.RemoveRange(startMenuRecentProgramsMaxCount, startMenuRecentPrograms.Count - startMenuRecentProgramsMaxCount);
            }
            RefreshStartMenuRecentProgramsSection();
            ApplyIconSizeSettings();

            if (persist)
            {
                SaveTaskbarSettings();
            }
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
            var oldSpotifyPreviousImage = btnSpotifyPrevious.Image;
            btnSpotifyPrevious.Image = ResizeIconImage(spotifyPreviousIconImage, GetSpotifyControlIconSize());
            oldSpotifyPreviousImage?.Dispose();
            var oldSpotifyPlayImage = btnSpotifyPlayPause.Image;
            btnSpotifyPlayPause.Image = ResizeIconImage(spotifyPlayIconImage, GetSpotifyControlIconSize());
            oldSpotifyPlayImage?.Dispose();
            var oldSpotifyNextImage = btnSpotifyNext.Image;
            btnSpotifyNext.Image = ResizeIconImage(spotifyNextIconImage, GetSpotifyControlIconSize());
            oldSpotifyNextImage?.Dispose();
            ApplyQuickLaunchButtonStyle(btnSpotifyPrevious);
            ApplyQuickLaunchButtonStyle(btnSpotifyPlayPause);
            ApplyQuickLaunchButtonStyle(btnSpotifyNext);

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

        private void OpenContextWindowProcessLocation()
        {
            var processId = GetProcessIdForWindow(contextMenuWindowHandle);
            var processPath = processId > 0 ? GetProcessPath(processId) : string.Empty;
            if (string.IsNullOrWhiteSpace(processPath) || !File.Exists(processPath))
            {
                System.Media.SystemSounds.Beep.Play();
                return;
            }

            LaunchProcess("explorer.exe", $"/select,\"{processPath}\"");
        }

        private void CopyContextWindowInfo()
        {
            if (!windowsByHandle.TryGetValue(contextMenuWindowHandle, out var window))
            {
                return;
            }

            var processId = GetProcessIdForWindow(contextMenuWindowHandle);
            var processPath = processId > 0 ? GetProcessPath(processId) : string.Empty;
            var processName = processId > 0 ? GetProcessName(processId) : string.Empty;
            Clipboard.SetText(string.Join(Environment.NewLine,
                $"Title: {window.Title}",
                $"Handle: 0x{contextMenuWindowHandle.ToInt64():X}",
                $"Process: {processName}",
                $"PID: {(processId > 0 ? processId.ToString(CultureInfo.InvariantCulture) : string.Empty)}",
                $"Path: {processPath}",
                $"State: {window.State}"));
        }

        private void EndContextWindowProcess()
        {
            if (!windowsByHandle.TryGetValue(contextMenuWindowHandle, out var window))
            {
                return;
            }

            var processId = GetProcessIdForWindow(contextMenuWindowHandle);
            if (processId <= 0 || processId == Environment.ProcessId)
            {
                return;
            }

            var processName = GetProcessName(processId);
            var processPath = GetProcessPath(processId);
            var result = MessageBox.Show(
                $"Ending a process from the taskbar can close more than this one window.{Environment.NewLine}{Environment.NewLine}" +
                $"{window.Title}{Environment.NewLine}" +
                $"{processName} (PID {processId}){Environment.NewLine}" +
                $"{processPath}{Environment.NewLine}{Environment.NewLine}" +
                "End this process?",
                "Taskbar Warning",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button2);
            if (result != DialogResult.Yes)
            {
                return;
            }

            try
            {
                using var process = Process.GetProcessById(processId);
                process.Kill();
                process.WaitForExit(500);
                RefreshTaskbar();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Unable to end process.{Environment.NewLine}{ex.Message}", "Taskbar", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static int GetProcessIdForWindow(IntPtr handle)
        {
            if (handle == IntPtr.Zero)
            {
                return 0;
            }

            WinApi.GetWindowThreadProcessId(handle, out var processId);
            return processId > int.MaxValue ? 0 : (int)processId;
        }

        private static string GetProcessPath(int processId)
        {
            try
            {
                using var process = Process.GetProcessById(processId);
                return process.MainModule?.FileName ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string GetProcessName(int processId)
        {
            try
            {
                using var process = Process.GetProcessById(processId);
                return process.ProcessName.EndsWith(".exe", StringComparison.OrdinalIgnoreCase)
                    ? process.ProcessName
                    : process.ProcessName + ".exe";
            }
            catch
            {
                return string.Empty;
            }
        }

        private void BtnStart_Click(object? sender, EventArgs e)
        {
            startMenuController.OnStartButtonClick();
        }

        private void BtnStart_MouseDown(object? sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left)
            {
                return;
            }

            startMenuController.OnStartButtonMouseDown();
        }

        private void ToggleStartMenu()
        {
            if (startMenu.Visible)
            {
                startMenu.Hide();
                ResetStartMenuSearch();
                ResetStartMenuSubmenuSessionCache();
                btnStart.Capture = false;
                ActiveControl = null;
                btnStart.Invalidate();
                btnStart.Update();
                return;
            }

            ResetStartMenuSearch();
            ResetStartMenuSubmenuSessionCache();

            startMenu.Show(btnStart, new Point(1, -Math.Max(startMenu.GetPreferredSize(Size.Empty).Height, 1)));
            RepositionVisibleStartMenu();
            var firstItem = startMenu.Items.OfType<ToolStripItem>().FirstOrDefault(IsNavigableTopLevelMenuItem);
            firstItem?.Select();
            btnStart.Invalidate();
        }

        private void CloseStartMenuIfVisible()
        {
            if (startMenu.Visible)
            {
                startMenu.Hide();
                ResetStartMenuSearch();
                ResetStartMenuSubmenuSessionCache();
                btnStart.Capture = false;
                ActiveControl = null;
                btnStart.Invalidate();
                btnStart.Update();
            }
        }

        private void RepositionVisibleStartMenu()
        {
            if (!startMenu.Visible)
            {
                return;
            }

            var startButtonLocation = btnStart.PointToScreen(Point.Empty);
            var menuX = startButtonLocation.X - 1;
            var menuHeight = Math.Max(startMenu.GetPreferredSize(Size.Empty).Height, 1);
            var menuY = startButtonLocation.Y - menuHeight - 4;
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

        private void InstallMouseHook()
        {
            if (mouseHookHandle != IntPtr.Zero)
            {
                return;
            }

            mouseHookHandle = WinApi.SetWindowsHookEx(WinApi.WH_MOUSE_LL, mouseProc, IntPtr.Zero, 0);
            if (mouseHookHandle == IntPtr.Zero)
            {
                Debug.WriteLine("Failed to install mouse hook.");
            }
        }

        private void UninstallMouseHook()
        {
            if (mouseHookHandle == IntPtr.Zero)
            {
                return;
            }

            WinApi.UnhookWindowsHookEx(mouseHookHandle);
            mouseHookHandle = IntPtr.Zero;
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
                trayIconFirstSeenUtc.Clear();
                trayInitialSyncComplete = false;
                trayIconsPanel.Controls.Clear();
                ResizeTrayPanel(allowShrink: true);
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
            var now = DateTime.UtcNow;
            var visibleCandidates = currentIcons.Where(icon => !icon.IsHidden && !ShouldHideTrayIcon(icon)).ToList();
            var visibleCandidateSet = visibleCandidates.ToHashSet();

            foreach (var icon in trayIconFirstSeenUtc.Keys.Where(icon => !visibleCandidateSet.Contains(icon)).ToList())
            {
                trayIconFirstSeenUtc.Remove(icon);
            }

            foreach (var icon in visibleCandidates)
            {
                trayIconFirstSeenUtc.TryAdd(icon, now);
            }

            var hasPendingIcons = false;
            var visibleIcons = visibleCandidates
                .Where(icon =>
                {
                    if (!trayInitialSyncComplete || trayIconControls.ContainsKey(icon))
                    {
                        return true;
                    }

                    if (!trayController.HasStableIdentity(icon))
                    {
                        hasPendingIcons = true;
                        return false;
                    }

                    if (trayIconFirstSeenUtc.TryGetValue(icon, out var firstSeenUtc) &&
                        (now - firstSeenUtc).TotalMilliseconds >= TrayIconNewIconStabilityMs)
                    {
                        return true;
                    }

                    hasPendingIcons = true;
                    return false;
                })
                .ToList();

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

            if (!trayIconsPanel.Controls.Contains(languageIndicatorLabel))
            {
                trayIconsPanel.Controls.Add(languageIndicatorLabel);
            }

            trayIconsPanel.Controls.SetChildIndex(customVolumeIcon, trayIconsPanel.Controls.Count - 1);
            trayIconsPanel.Controls.SetChildIndex(languageIndicatorLabel, trayIconsPanel.Controls.Count - 1);
            UpdateInputMethodIndicators();

            ResizeTrayPanel();
            RefreshTrayHostPlacement();
            trayInitialSyncComplete = true;

            if (hasPendingIcons && !trayIconStabilityTimer.Enabled)
            {
                trayIconStabilityTimer.Start();
            }
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
            var image = ConvertImageSourceToBitmap(icon.Icon, GetScaledSmallIconSize());
            if (image == null)
            {
                image = CreateTransparentPlaceholderIcon(GetScaledSmallIconSize());
            }

            var oldImage = control.Image;
            control.Image = image;
            oldImage?.Dispose();
        }

        private static Bitmap CreateTransparentPlaceholderIcon(int size)
        {
            var clampedSize = Math.Max(1, size);
            var bitmap = new Bitmap(clampedSize, clampedSize, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using var graphics = Graphics.FromImage(bitmap);
            graphics.Clear(Color.Transparent);
            return bitmap;
        }

        private void UpdateInputMethodIndicators()
        {
            var imeTargetWindow = GetImeTargetWindowHandle();
            var threadId = imeTargetWindow == IntPtr.Zero
                ? 0u
                : WinApi.GetWindowThreadProcessId(imeTargetWindow, out _);
            var keyboardLayout = WinApi.GetKeyboardLayout(threadId);
            var culture = GetCultureFromKeyboardLayout(keyboardLayout);

            languageIndicatorLabel.Text = GetLanguageIndicatorText(culture);

            var imeOpen = IsImeOpenForWindow(imeTargetWindow);
            imeIndicatorLabel.Text = imeOpen ? "Im" : "Al";
        }

        private IntPtr GetImeTargetWindowHandle()
        {
            var foregroundWindow = WinApi.GetForegroundWindow();
            if (foregroundWindow != IntPtr.Zero && foregroundWindow != Handle && !IsDesktopOrShellWindow(foregroundWindow))
            {
                return foregroundWindow;
            }

            if (interactionStateMachine.ActiveWindowHandle != IntPtr.Zero && interactionStateMachine.ActiveWindowHandle != Handle)
            {
                return interactionStateMachine.ActiveWindowHandle;
            }

            if (interactionStateMachine.ForegroundWindowBeforeTaskClick != IntPtr.Zero && interactionStateMachine.ForegroundWindowBeforeTaskClick != Handle)
            {
                return interactionStateMachine.ForegroundWindowBeforeTaskClick;
            }

            return foregroundWindow;
        }

        private static CultureInfo? GetCultureFromKeyboardLayout(IntPtr keyboardLayout)
        {
            var langId = (int)((long)keyboardLayout & 0xFFFF);
            if (langId <= 0)
            {
                return null;
            }

            try
            {
                return CultureInfo.GetCultureInfo(langId);
            }
            catch (CultureNotFoundException)
            {
                return null;
            }
        }

        private static string GetLanguageIndicatorText(CultureInfo? culture)
        {
            if (culture is null)
            {
                return "--";
            }

            var code = culture.TwoLetterISOLanguageName;
            if (string.IsNullOrWhiteSpace(code) || code.Equals("iv", StringComparison.OrdinalIgnoreCase))
            {
                return "--";
            }

            if (code.Length > 2)
            {
                code = code[..2];
            }

            return char.ToUpperInvariant(code[0]) + code[1..].ToLowerInvariant();
        }

        private static bool IsImeOpenForWindow(IntPtr windowHandle)
        {
            return TryGetImeStatus(windowHandle, out var imeOpen, out _, out _) && imeOpen;
        }

        private static bool TryGetImeStatus(IntPtr windowHandle, out bool imeOpen, out int conversionMode, out int sentenceMode)
        {
            imeOpen = false;
            conversionMode = 0;
            sentenceMode = 0;

            if (windowHandle == IntPtr.Zero)
            {
                return false;
            }

            var inputContext = WinApi.ImmGetContext(windowHandle);
            if (inputContext == IntPtr.Zero)
            {
                return false;
            }

            try
            {
                imeOpen = WinApi.ImmGetOpenStatus(inputContext);
                var gotConversion = WinApi.ImmGetConversionStatus(inputContext, out conversionMode, out sentenceMode);
                if (!gotConversion)
                {
                    conversionMode = 0;
                    sentenceMode = 0;
                }

                return true;
            }
            finally
            {
                WinApi.ImmReleaseContext(windowHandle, inputContext);
            }
        }

        private void ApplyImeStateToForegroundWindow(bool imeOpen, int conversionMode, int sentenceMode)
        {
            var foregroundWindow = WinApi.GetForegroundWindow();
            if (foregroundWindow == IntPtr.Zero)
            {
                return;
            }

            var inputContext = WinApi.ImmGetContext(foregroundWindow);
            if (inputContext == IntPtr.Zero)
            {
                return;
            }

            try
            {
                WinApi.ImmSetOpenStatus(inputContext, imeOpen);
                WinApi.ImmSetConversionStatus(inputContext, conversionMode, sentenceMode);
            }
            finally
            {
                WinApi.ImmReleaseContext(foregroundWindow, inputContext);
            }

            UpdateInputMethodIndicators();
        }

        private void ResizeTrayPanel(bool allowShrink = false)
        {
            var desiredCount = trayIconControls.Count + 2;
            if (desiredCount >= trayIconSlotCount || allowShrink)
            {
                trayIconSlotCount = desiredCount;
                trayResizeShrinkTimer.Stop();
            }
            else if (!trayResizeShrinkTimer.Enabled)
            {
                trayResizeShrinkTimer.Start();
            }

            var count = Math.Max(desiredCount, trayIconSlotCount);
            var trayRows = GetTrayIconRows();
            var trayColumns = Math.Max(1, (int)Math.Ceiling(count / (double)trayRows));
            trayIconsPanel.FlowDirection = trayRows > 1 ? FlowDirection.TopDown : FlowDirection.LeftToRight;
            trayIconsPanel.WrapContents = trayRows > 1;
            trayIconsPanel.Padding = trayRows > 1
                ? new Padding(2, Math.Max(1, (trayIconsPanel.ClientSize.Height - (trayRows * 18)) / 2), 2, 0)
                : new Padding(2, 4, 2, 0);
            trayIconsPanel.Width = Math.Clamp((trayColumns * TrayIconCellSize) + trayIconsPanel.Padding.Horizontal + 4, 32, 220);
            notificationAreaPanel.Width = trayIconsPanel.Width + lblClock.Width + notificationAreaPanel.Padding.Horizontal + 6;
            trayIconsPanel.PerformLayout();
            notificationAreaPanel.Invalidate();
        }

        private int GetTrayIconRows()
        {
            if (taskbarRows <= 1)
            {
                return 1;
            }

            return Math.Clamp(taskbarRows + 1, 2, 4);
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
            trayIconInteractionService.ForwardMouseDown(icon, e.Button, GetCursorPositionParam(), SystemInformation.DoubleClickTime);
        }

        private void TrayIcon_MouseUp(object? sender, MouseEventArgs e)
        {
            if (sender is not PictureBox control || control.Tag is not TrayNotifyIcon icon)
            {
                return;
            }

            if (e.Button == MouseButtons.Left && IsVolumeTrayIcon(icon))
            {
                if (useClassicVolumePopup)
                {
                    ShowClassicVolumePopup();
                }
                else
                {
                    OpenSystemVolumeMixer();
                }

                return;
            }

            UpdateTrayIconPlacement(icon, control);
            trayIconInteractionService.ForwardMouseUp(icon, e.Button, GetCursorPositionParam(), SystemInformation.DoubleClickTime);
        }

        private void TrayIcon_MouseDoubleClick(object? sender, MouseEventArgs e)
        {
            if (sender is not PictureBox control || control.Tag is not TrayNotifyIcon icon)
            {
                return;
            }

            if (e.Button == MouseButtons.Left && trayController.ShouldOpenWindowsSecurityCenterOnDoubleClick(icon))
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

            if (useClassicVolumePopup)
            {
                ShowClassicVolumePopup();
                return;
            }

            OpenSystemVolumeMixer();
        }

        private void TrayIcon_MouseMove(object? sender, MouseEventArgs e)
        {
            if (sender is not PictureBox control || control.Tag is not TrayNotifyIcon icon)
            {
                return;
            }

            if (trayController.ShouldSuppressHoverForwarding(icon))
            {
                return;
            }

            trayIconInteractionService.ForwardMouseMove(icon, GetCursorPositionParam());
        }

        private void TrayIcon_MouseEnter(object? sender, EventArgs e)
        {
            if (sender is not PictureBox control || control.Tag is not TrayNotifyIcon icon)
            {
                return;
            }

            if (trayController.ShouldSuppressHoverForwarding(icon))
            {
                return;
            }

            UpdateTrayIconPlacement(icon, control);
            trayIconInteractionService.ForwardMouseEnter(icon, GetCursorPositionParam());
        }

        private void TrayIcon_MouseLeave(object? sender, EventArgs e)
        {
            if (sender is not PictureBox control || control.Tag is not TrayNotifyIcon icon)
            {
                return;
            }

            if (trayController.ShouldSuppressHoverForwarding(icon))
            {
                return;
            }

            trayIconInteractionService.ForwardMouseLeave(icon, GetCursorPositionParam());
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
            return trayController.IsVolumeTrayIcon(icon);
        }

        private bool ShouldHideTrayIcon(TrayNotifyIcon icon)
        {
            return trayController.ShouldHideTrayIcon(icon);
        }

        private string GetTrayHoverMessage(TrayNotifyIcon icon)
        {
            return trayController.GetHoverMessage(icon);
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

        private void OpenSystemVolumeMixer()
        {
            try
            {
                if (volumePopup.Visible)
                {
                    volumePopup.Hide();
                }

                var process = Process.Start(new ProcessStartInfo
                {
                    FileName = "sndvol.exe",
                    UseShellExecute = true
                });

                PositionProcessWindowNearTray(process);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to open system volume mixer: {ex.Message}");
            }
        }

        private void ShowClassicVolumePopup()
        {
            var iconBounds = customVolumeIcon.RectangleToScreen(customVolumeIcon.ClientRectangle);
            var popupX = iconBounds.Right - volumePopup.Width;
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

        private static void PositionProcessWindowNearTray(Process? process)
        {
            var targetProcessId = process?.Id ?? 0;
            var windowHandle = IntPtr.Zero;

            for (var attempt = 0; attempt < 40; attempt++)
            {
                if (targetProcessId > 0)
                {
                    windowHandle = FindTopLevelWindowByProcessId((uint)targetProcessId);
                }

                if (windowHandle == IntPtr.Zero)
                {
                    var foregroundWindow = WinApi.GetForegroundWindow();
                    if (foregroundWindow != IntPtr.Zero)
                    {
                        WinApi.GetWindowThreadProcessId(foregroundWindow, out var foregroundPid);
                        if (foregroundPid != 0 && IsSndVolProcess(foregroundPid))
                        {
                            windowHandle = foregroundWindow;
                        }
                    }
                }

                if (windowHandle != IntPtr.Zero)
                {
                    break;
                }

                Thread.Sleep(50);
            }

            if (windowHandle == IntPtr.Zero || !WinApi.GetWindowRect(windowHandle, out var rect))
            {
                return;
            }

            var width = Math.Max(1, rect.Right - rect.Left);
            var height = Math.Max(1, rect.Bottom - rect.Top);
            var workArea = Screen.PrimaryScreen?.WorkingArea ?? new Rectangle(0, 0, 1024, 768);
            var trayLift = Math.Max(12, (int)Math.Round(14 * (Form.ActiveForm?.DeviceDpi ?? 96) / 96.0));
            var targetX = Math.Max(workArea.Left, workArea.Right - width - 8);
            var targetY = Math.Max(workArea.Top, workArea.Bottom - height - trayLift);
            WinApi.MoveWindow(windowHandle, targetX, targetY, width, height, true);
        }

        private static IntPtr FindTopLevelWindowByProcessId(uint processId)
        {
            var foundWindow = IntPtr.Zero;
            WinApi.EnumWindows((hWnd, _) =>
            {
                if (!WinApi.IsWindowVisible(hWnd))
                {
                    return true;
                }

                WinApi.GetWindowThreadProcessId(hWnd, out var windowProcessId);
                if (windowProcessId == processId)
                {
                    foundWindow = hWnd;
                    return false;
                }

                return true;
            }, IntPtr.Zero);

            return foundWindow;
        }

        private static bool IsSndVolProcess(uint processId)
        {
            try
            {
                var process = Process.GetProcessById((int)processId);
                return process.ProcessName.Equals("SndVol", StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
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
                var isInjected = (keyInfo.flags & WinApi.LLKHF_INJECTED) != 0;

                if (message == WinApi.WM_KEYDOWN || message == WinApi.WM_SYSKEYDOWN)
                {
                    if (isInjected)
                    {
                        return WinApi.CallNextHookEx(keyboardHookHandle, nCode, wParam, lParam);
                    }

                    if (vkCode == (int)Keys.F4 && IsAltKeyDown() && WinApi.GetForegroundWindow() == Handle)
                    {
                        return (IntPtr)1;
                    }

                    if (isWinKey)
                    {
                        activeWinKeyVkCode = vkCode;
                        winKeyPressed = true;
                        nonWinKeyPressedWhileWinHeld = false;
                        injectedWinDownForChord = false;
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
                        if (!injectedWinDownForChord)
                        {
                            InjectWinKey(activeWinKeyVkCode, keyUp: false);
                            injectedWinDownForChord = true;
                        }

                        // Let other Win+ combinations pass through to the system.
                        nonWinKeyPressedWhileWinHeld = true;
                        return WinApi.CallNextHookEx(keyboardHookHandle, nCode, wParam, lParam);
                    }

                    if (startMenu.Visible)
                    {
                        if (IsStartMenuSearchTextBoxFocused() && !IsStartMenuNavigationVirtualKey(vkCode))
                        {
                            return WinApi.CallNextHookEx(keyboardHookHandle, nCode, wParam, lParam);
                        }

                        if (HandleStartMenuSearchKey(vkCode))
                        {
                            return (IntPtr)1;
                        }
                    }
                }
                else if ((message == WinApi.WM_KEYUP || message == WinApi.WM_SYSKEYUP) && isWinKey)
                {
                    if (isInjected)
                    {
                        return WinApi.CallNextHookEx(keyboardHookHandle, nCode, wParam, lParam);
                    }

                    var shouldToggleStart = winKeyPressed && !nonWinKeyPressedWhileWinHeld;
                    if (injectedWinDownForChord)
                    {
                        InjectWinKey(activeWinKeyVkCode, keyUp: true);
                    }

                    winKeyPressed = false;
                    nonWinKeyPressedWhileWinHeld = false;
                    injectedWinDownForChord = false;
                    activeWinKeyVkCode = 0;

                    if (shouldToggleStart && !IsDisposed)
                    {
                        BeginInvoke(new Action(ToggleStartMenu));
                    }

                    return (IntPtr)1;
                }
            }

            return WinApi.CallNextHookEx(keyboardHookHandle, nCode, wParam, lParam);
        }

        private static void InjectWinKey(int vkCode, bool keyUp)
        {
            if (vkCode != WinApi.VK_LWIN && vkCode != WinApi.VK_RWIN)
            {
                return;
            }

            var flags = WinApi.KEYEVENTF_EXTENDEDKEY;
            if (keyUp)
            {
                flags |= WinApi.KEYEVENTF_KEYUP;
            }

            WinApi.keybd_event((byte)vkCode, 0, flags, UIntPtr.Zero);
        }

        private IntPtr MouseHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                var message = wParam.ToInt32();
                if (message == WinApi.WM_LBUTTONDOWN || message == WinApi.WM_RBUTTONDOWN || message == WinApi.WM_MBUTTONDOWN)
                {
                    var mouseInfo = Marshal.PtrToStructure<WinApi.MSLLHOOKSTRUCT>(lParam);
                    var clickPoint = mouseInfo.pt;

                    if (startMenu.Visible && !IsPointInStartMenuOrDropDowns(clickPoint))
                    {
                        BeginInvoke(new Action(CloseStartMenuIfVisible));
                    }

                    if (IsTaskbarContextMenuOpen() && !IsPointInContextMenuOrDropDowns(taskbarContextMenu, clickPoint))
                    {
                        BeginInvoke(new Action(CloseTaskbarContextMenuTree));
                    }
                }
            }

            return WinApi.CallNextHookEx(mouseHookHandle, nCode, wParam, lParam);
        }

        private bool IsTaskbarContextMenuOpen()
        {
            return taskbarContextMenu.Visible || HasVisibleDropDowns(taskbarContextMenu);
        }

        private void CloseTaskbarContextMenuTree()
        {
            foreach (var menuItem in taskbarContextMenu.Items.OfType<ToolStripMenuItem>())
            {
                HideDropDownTree(menuItem);
            }

            taskbarContextMenu.Close(ToolStripDropDownCloseReason.AppClicked);
        }

        private bool IsPointInStartMenuOrDropDowns(Point point)
        {
            if (IsPointInContextMenuOrDropDowns(startMenu, point))
                return true;

            // Check taskbar (clicking on taskbar should not close start menu)
            var taskbarBounds = RectangleToScreen(ClientRectangle);
            if (taskbarBounds.Contains(point))
                return true;

            return false;
        }

        private static bool IsPointInContextMenuOrDropDowns(ContextMenuStrip menu, Point point)
        {
            if (!menu.Visible)
            {
                return false;
            }

            var menuBounds = new Rectangle(menu.Location, menu.Size);
            if (menuBounds.Contains(point))
            {
                return true;
            }

            return HasNestedDropDowns(menu, point);
        }

        private static bool HasNestedDropDowns(ToolStripDropDown menu, Point point)
        {
            foreach (ToolStripItem item in menu.Items)
            {
                if (item is ToolStripMenuItem menuItem && menuItem.DropDown is { Visible: true } dropDown)
                {
                    var dropDownBounds = new Rectangle(dropDown.Location, dropDown.Size);
                    if (dropDownBounds.Contains(point))
                        return true;

                    if (HasNestedDropDowns(dropDown, point))
                        return true;
                }
            }
            return false;
        }

        private static bool HasVisibleDropDowns(ToolStripDropDown menu)
        {
            foreach (ToolStripItem item in menu.Items)
            {
                if (item is not ToolStripMenuItem menuItem || menuItem.DropDown is not { Visible: true } dropDown)
                {
                    continue;
                }

                return true;
            }

            return false;
        }

        private static void HideDropDownTree(ToolStripMenuItem menuItem)
        {
            foreach (var child in menuItem.DropDownItems.OfType<ToolStripMenuItem>())
            {
                HideDropDownTree(child);
            }

            if (menuItem.DropDown.Visible)
            {
                menuItem.HideDropDown();
            }
        }

        private void BringExplorerToFront()
        {
            var explorer = new Form1();
            explorer.Show();
            explorer.BringToFront();
            explorer.Activate();
        }

        private void UpdateSpotifyToolbarState()
        {
            var now = DateTime.UtcNow;
            if ((now - spotifyStateLastCheckedUtc).TotalSeconds < 1)
            {
                return;
            }

            spotifyStateLastCheckedUtc = now;
            var running = spotifyController.IsSpotifyRunning();
            if (running == isSpotifyRunning)
            {
                return;
            }

            isSpotifyRunning = running;
            btnSpotifyPrevious.Enabled = true;
            btnSpotifyPlayPause.Enabled = true;
            btnSpotifyNext.Enabled = true;
        }

        private void UpdateSpotifyToolbarLayout()
        {
            var showTrackText = showSpotifyToolbar && taskbarRows > 1;
            var buttonWidth = (QuickLaunchButtonSize * 3) + (TaskbarButtonGap * 2) + spotifyToolbarPanel.Padding.Horizontal;
            spotifyNowPlayingLabel.Visible = showTrackText;
            spotifyControlsPanel.Height = QuickLaunchButtonSize;
            spotifyToolbarPanel.Width = showTrackText ? Math.Max(180, buttonWidth) : buttonWidth;

            if (!showTrackText)
            {
                spotifyNowPlayingLabel.Text = string.Empty;
                spotifyToolbarPanel.Height = QuickLaunchButtonSize + spotifyToolbarPanel.Padding.Vertical;
                return;
            }

            spotifyToolbarPanel.Height = Math.Max(QuickLaunchButtonSize * 2, TaskbarRowHeight * Math.Min(taskbarRows, 2));
            var nowPlaying = spotifyController.GetNowPlayingTitle();
            spotifyNowPlayingLabel.Text = string.IsNullOrWhiteSpace(nowPlaying)
                ? (isSpotifyRunning ? "Spotify" : "Spotify is not running")
                : nowPlaying;
            quickLaunchToolTip.SetToolTip(spotifyNowPlayingLabel, spotifyNowPlayingLabel.Text);
        }

        private void RefreshTaskbar()
        {
            DetectFullscreenWindow();
            SetTaskbarBounds();
            UpdateClockDisplay();
            UpdateInputMethodIndicators();
            SyncVolumeToolbarFromSystem();
            UpdateSpotifyToolbarState();
            UpdateSpotifyToolbarLayout();

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
                interactionStateMachine.SetActiveWindowHandle(foregroundHandle);
            }
            else
            {
                var activeByState = windows
                    .FirstOrDefault(w => w.State == ApplicationWindow.WindowState.Active)?.Handle ?? IntPtr.Zero;

                if (currentHandles.Contains(activeByState))
                {
                    interactionStateMachine.SetActiveWindowHandle(activeByState);
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
                windowsMinimizedByShowDesktop.Remove(handle);
            }

            ResizeTaskButtons();
        }

        private void UpdateClockDisplay()
        {
            var now = DateTime.Now;
            clockLineDate = now.ToString("yyyy/MM/dd");
            clockLineWeekday = now.ToString("dddd");
            clockLineTime = now.ToString("tt hh:mm");
            lblClock.Text = string.Empty;

            var clockLines = GetClockLinesForRendering();
            var lineWidth = clockLines
                .Select(line => TextRenderer.MeasureText(line, lblClock.Font).Width)
                .DefaultIfEmpty(0)
                .Max();
            var preferredWidth = lineWidth + 14;
            lblClock.Width = Math.Clamp(preferredWidth, 110, 260);
            lblClock.Invalidate();
        }

        private void LblClock_Paint(object? sender, PaintEventArgs e)
        {
            var lines = GetClockLinesForRendering();
            if (lines.Count == 0)
            {
                return;
            }

            var lineHeight = TextRenderer.MeasureText("Hg", lblClock.Font).Height;
            var totalHeight = (lineHeight * lines.Count) + (ClockLineSpacingPx * Math.Max(0, lines.Count - 1));
            var top = Math.Max(0, (lblClock.ClientSize.Height - totalHeight) / 2);
            var format = TextFormatFlags.HorizontalCenter | TextFormatFlags.NoPadding | TextFormatFlags.EndEllipsis;

            for (var i = 0; i < lines.Count; i++)
            {
                var y = top + (i * (lineHeight + ClockLineSpacingPx));
                var rect = new Rectangle(0, y, lblClock.ClientSize.Width, lineHeight);
                TextRenderer.DrawText(e.Graphics, lines[i], lblClock.Font, rect, lblClock.ForeColor, format);
            }
        }

        private List<string> GetClockLinesForRendering()
        {
            var lineHeight = TextRenderer.MeasureText("Hg", lblClock.Font).Height;
            var maxLines = (lblClock.ClientSize.Height + ClockLineSpacingPx) / (lineHeight + ClockLineSpacingPx);

            return maxLines switch
            {
                >= 3 => new List<string> { clockLineDate, clockLineWeekday, clockLineTime },
                2 => new List<string> { clockLineDate, clockLineTime },
                1 => new List<string> { clockLineTime },
                _ => new List<string>()
            };
        }

        private void DisposeTaskButton(Button button, bool removeFromPanel)
        {
            button.Click -= TaskButton_Click;
            button.MouseDown -= TaskButton_MouseDown;
            button.MouseUp -= TaskButton_MouseUp;
            button.MouseEnter -= TaskButton_MouseEnter;
            button.MouseMove -= TaskButton_MouseMove;
            button.MouseLeave -= TaskButton_MouseLeave;

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
                Font = normalFont,
                ForeColor = taskbarFontColor
            };
            ApplyTaskbarButtonStyle(button);
            AttachNoFocusCue(button);

            button.Click += TaskButton_Click;
            button.MouseDown += TaskButton_MouseDown;
            button.MouseUp += TaskButton_MouseUp;
            button.MouseEnter += TaskButton_MouseEnter;
            button.MouseMove += TaskButton_MouseMove;
            button.MouseLeave += TaskButton_MouseLeave;
            return button;
        }

        private void UpdateTaskButton(Button button, ApplicationWindow window)
        {
            var title = string.IsNullOrWhiteSpace(window.Title) ? "Untitled" : window.Title.Trim();
            button.AccessibleName = title;
            taskToolTip.SetToolTip(button, string.Empty);

            var taskButtonIconSize = GetScaledTaskButtonIconSize();
            if (button.Image == null || button.Image.Width != taskButtonIconSize || button.Image.Height != taskButtonIconSize)
            {
                var image = ConvertImageSourceToBitmap(window.Icon, taskButtonIconSize);
                if (image != null)
                {
                    var oldImage = button.Image;
                    button.Image = AddIconMargin(image, 2, 1);
                    oldImage?.Dispose();
                }
            }

            var isActive = window.Handle == interactionStateMachine.ActiveWindowHandle || window.State == ApplicationWindow.WindowState.Active;

            if (isActive)
            {
                button.Font = activeFont;
                button.BackColor = taskbarButtonStyle == TaskbarButtonStyle.Win98
                    ? Win98ActiveTaskButtonBackColor
                    : taskbarLightColor;
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
            var imageSpace = button.Image != null ? button.Image.Width + 8 : 0;
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
                var scaled = new Bitmap(iconSize, iconSize, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                using var graphics = Graphics.FromImage(scaled);
                graphics.Clear(Color.Transparent);
                graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                graphics.SmoothingMode = SmoothingMode.HighQuality;
                graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
                graphics.CompositingQuality = CompositingQuality.HighQuality;
                graphics.DrawImage(image, new Rectangle(0, 0, iconSize, iconSize), new Rectangle(0, 0, image.Width, image.Height), GraphicsUnit.Pixel);
                return scaled;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to convert task icon: {ex.Message}");
                return null;
            }
        }

        private int GetScaledTaskButtonIconSize()
        {
            var scale = DeviceDpi <= 0 ? 1.0 : DeviceDpi / 96.0;
            var scaled = (int)Math.Round(taskIconSize * scale);
            var maxIconSize = Math.Max(16, TaskbarButtonHeight - 4);
            return Math.Clamp(scaled, 16, maxIconSize);
        }

        private static Bitmap ResizeIconImage(Image source, int size)
        {
            var scaled = new Bitmap(size, size, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using var graphics = Graphics.FromImage(scaled);
            graphics.Clear(Color.Transparent);
            graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
            graphics.SmoothingMode = SmoothingMode.HighQuality;
            graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
            graphics.CompositingQuality = CompositingQuality.HighQuality;
            graphics.DrawImage(source, new Rectangle(0, 0, size, size), new Rectangle(0, 0, source.Width, source.Height), GraphicsUnit.Pixel);
            return scaled;
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

        private int GetSpotifyControlIconSize()
        {
            var scale = DeviceDpi <= 0 ? 1.0 : DeviceDpi / 96.0;
            var scaled = (int)Math.Round(18 * scale);
            return Math.Clamp(scaled, 18, Math.Max(18, TaskbarButtonHeight - 4));
        }

        private void ResizeTaskButtons()
        {
            var buttonCount = taskButtonsPanel.Controls.Count;
            if (buttonCount == 0)
            {
                return;
            }

            var rows = GetTaskButtonRows();
            var columns = Math.Max(1, (int)Math.Ceiling(buttonCount / (double)rows));
            var totalGapWidth = Math.Max(0, columns - 1) * TaskbarButtonGap;
            var availableWidth = Math.Max(taskButtonsPanel.ClientSize.Width - totalGapWidth - taskButtonsPanel.Padding.Horizontal, btnStart.Width);

            var minWidth = rows > 1 ? TaskButtonMinWidthMultiRow : btnStart.Width;
            var widthPerButton = Math.Clamp(availableWidth / columns, minWidth, TaskButtonMaxWidth);

            taskButtonsPanel.WrapContents = rows > 1;
            taskButtonsPanel.FlowDirection = FlowDirection.LeftToRight;

            foreach (Control control in taskButtonsPanel.Controls)
            {
                control.Width = widthPerButton;
                control.Height = TaskbarButtonHeight;
                control.Margin = new Padding(0, 0, TaskbarButtonGap, TaskButtonRowBottomGap);
            }

            taskButtonsPanel.PerformLayout();
        }

        private void RefreshTaskButtonAndTrayLayout()
        {
            ApplyToolbarLayout(refreshBounds: false);
            ResizeTaskButtons();
            ResizeTrayPanel();
            RefreshTrayHostPlacement();
        }

        private static int GetMasterVolumePercent()
        {
            if (!TryGetAudioEndpointVolume(out var endpointVolume) || endpointVolume is null)
            {
                return 50;
            }
            try
            {
                _ = endpointVolume.GetMasterVolumeLevelScalar(out var level);
                return Math.Clamp((int)Math.Round(level * 100f), 0, 100);
            }
            catch
            {
                return 50;
            }
            finally
            {
                if (endpointVolume is not null)
                {
                    Marshal.ReleaseComObject(endpointVolume);
                }
            }
        }

        private static void SetMasterVolumePercent(int value)
        {
            if (!TryGetAudioEndpointVolume(out var endpointVolume) || endpointVolume is null)
            {
                return;
            }
            try
            {
                var level = Math.Clamp(value, 0, 100) / 100f;
                _ = endpointVolume.SetMasterVolumeLevelScalar(level, Guid.Empty);
            }
            catch
            {
            }
            finally
            {
                if (endpointVolume is not null)
                {
                    Marshal.ReleaseComObject(endpointVolume);
                }
            }
        }

        private static bool TryGetAudioEndpointVolume(out IAudioEndpointVolume? endpointVolume)
        {
            endpointVolume = null;
            object? enumeratorObject = null;
            IMMDeviceEnumerator? deviceEnumerator = null;
            IMMDevice? device = null;

            try
            {
                enumeratorObject = Activator.CreateInstance(Type.GetTypeFromCLSID(CLSID_MMDeviceEnumerator)!);
                if (enumeratorObject is not IMMDeviceEnumerator enumerator)
                {
                    return false;
                }

                deviceEnumerator = enumerator;
                if (deviceEnumerator.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia, out device!) != 0 || device is null)
                {
                    return false;
                }

                var endpointVolumeGuid = IID_IAudioEndpointVolume;
                if (device.Activate(ref endpointVolumeGuid, CLSCTX_ALL, IntPtr.Zero, out var endpointObject) != 0)
                {
                    return false;
                }

                endpointVolume = endpointObject as IAudioEndpointVolume;
                return endpointVolume is not null;
            }
            catch
            {
                return false;
            }
            finally
            {
                if (device is not null)
                {
                    Marshal.ReleaseComObject(device);
                }

                if (deviceEnumerator is not null)
                {
                    Marshal.ReleaseComObject(deviceEnumerator);
                }

                if (enumeratorObject is not null && Marshal.IsComObject(enumeratorObject))
                {
                    Marshal.ReleaseComObject(enumeratorObject);
                }
            }
        }

        private void SyncVolumeToolbarFromSystem()
        {
            if (suppressVolumeToolbarEvents || isVolumeToolbarDragging || volumeToolbarTrackBar.IsDisposed)
            {
                return;
            }

            var volume = GetMasterVolumePercent();
            if (volumeToolbarTrackBar.Value == volume)
            {
                return;
            }

            suppressVolumeToolbarEvents = true;
            volumeToolbarTrackBar.Value = volume;
            suppressVolumeToolbarEvents = false;
        }

        private void VolumeToolbarTrackBar_ValueChanged(object? sender, EventArgs e)
        {
            if (suppressVolumeToolbarEvents)
            {
                return;
            }

            SetMasterVolumePercent(volumeToolbarTrackBar.Value);
        }

        private void VolumeToolbarTrackBar_MouseDown(object? sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                isVolumeToolbarDragging = true;
            }
        }

        private void VolumeToolbarTrackBar_MouseUp(object? sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left)
            {
                return;
            }

            isVolumeToolbarDragging = false;
            SetMasterVolumePercent(volumeToolbarTrackBar.Value);
        }

        private void RefreshTrayHostPlacement()
        {
            UpdateTrayHostSizeData();

            foreach (var pair in trayIconControls)
            {
                UpdateTrayIconPlacement(pair.Key, pair.Value);
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
                interactionStateMachine.SetActiveWindowHandle(handle);
            }
            else if (interactionStateMachine.ShouldMinimizeTaskWindow(handle))
            {
                window.Minimize();
                if (!window.IsMinimized && !WinApi.IsIconic(handle))
                {
                    WinApi.ShowWindow(handle, WinApi.SW_MINIMIZE);
                }

                interactionStateMachine.ClearActiveWindowHandleIfMatches(handle);
            }
            else
            {
                window.BringToFront();
                WinApi.SetForegroundWindow(handle);
                interactionStateMachine.SetActiveWindowHandle(handle);
            }

            interactionStateMachine.ClearTaskButtonMouseState();

            RefreshTaskbar();
        }

        private void TaskButton_MouseDown(object? sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left)
            {
                return;
            }

            interactionStateMachine.RecordTaskButtonMouseDown(WinApi.GetForegroundWindow());
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

        private void TaskButton_MouseEnter(object? sender, EventArgs e)
        {
            if (sender is not Button button)
            {
                return;
            }

            ShowTaskButtonToolTip(button);
        }

        private void TaskButton_MouseMove(object? sender, MouseEventArgs e)
        {
            if (sender is not Button button)
            {
                return;
            }

            ShowTaskButtonToolTip(button);
        }

        private void TaskButton_MouseLeave(object? sender, EventArgs e)
        {
            if (sender is Button button)
            {
                taskToolTip.Hide(button);
                taskToolTip.Hide(this);
            }
        }

        private void ShowTaskButtonToolTip(Button button)
        {
            var title = button.AccessibleName;
            if (string.IsNullOrWhiteSpace(title))
            {
                return;
            }

            // Anchor tooltip to task button top-center in form coordinates.
            var screenAnchor = button.PointToScreen(new Point(button.Width / 2, 0));
            var formAnchor = PointToClient(new Point(screenAnchor.X - 90, screenAnchor.Y - 28));
            taskToolTip.Show(title, this, formAnchor, 4000);
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
                    clockContextMenu.Visible || languageIndicatorContextMenu.Visible || imeIndicatorContextMenu.Visible || quickLaunchItemContextMenu.Visible ||
                    volumePopup.Visible || (timeDetailsPopup?.Visible ?? false) || ContainsFocus;

                taskbarY = keepVisible ? screenBounds.Bottom - totalHeight : screenBounds.Bottom - 2;
            }

            var newBounds = new Rectangle(screenBounds.Left, taskbarY, screenBounds.Width, totalHeight);

            if (isAppBarRegistered && !autoHideTaskbar)
            {
                UpdateAppBarPosition();
            }
            else if (Bounds != newBounds)
            {
                Bounds = newBounds;
                RefreshTrayHostPlacement();
            }
        }

        private void UpdateAppBarPosition()
        {
            if (!isAppBarRegistered)
            {
                return;
            }

            var screenBounds = Screen.PrimaryScreen?.Bounds ?? new Rectangle(0, 0, 1024, 768);
            var appBarData = new WinApi.APPBARDATA
            {
                cbSize = (uint)Marshal.SizeOf<WinApi.APPBARDATA>(),
                hWnd = Handle,
                uCallbackMessage = (uint)WinApi.WM_USER + 0x7FFF,
                uEdge = (uint)WinApi.AppBarEdge.Bottom,
                rc = new WinApi.RECT(
                    screenBounds.Left,
                    screenBounds.Bottom - TotalTaskbarHeight,
                    screenBounds.Right,
                    screenBounds.Bottom),
                lParam = IntPtr.Zero
            };

            WinApi.SHAppBarMessage((uint)WinApi.AppBarMsg.QueryPos, ref appBarData);
            appBarData.rc.Top = Math.Max(screenBounds.Top, appBarData.rc.Bottom - TotalTaskbarHeight);
            appBarData.rc.Left = screenBounds.Left;
            appBarData.rc.Right = screenBounds.Right;

            WinApi.SHAppBarMessage((uint)WinApi.AppBarMsg.SetPos, ref appBarData);
            var appBarBounds = appBarData.rc.ToRectangle();
            if (Bounds != appBarBounds)
            {
                Bounds = appBarBounds;
            }

            RefreshTrayHostPlacement();
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



