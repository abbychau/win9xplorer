using System.Drawing;

namespace win9xplorer
{
    internal sealed record TaskbarRuntimeSettings(
        int StartMenuIconSize,
        int TaskIconSize,
        bool LazyLoadProgramsSubmenu,
        bool PlayVolumeFeedbackSound,
        bool UseClassicVolumePopup,
        int StartMenuSubmenuOpenDelayMs,
        int StartMenuRecentProgramsMaxCount,
        bool AutoHideTaskbar,
        bool StartOnWindowsStartup,
        string TaskbarButtonStyle,
        string TaskbarFontName,
        float TaskbarFontSize,
        Color TaskbarFontColor,
        bool TaskbarLocked,
        int TaskbarRows,
        Color TaskbarBaseColor,
        Color TaskbarLightColor,
        Color TaskbarDarkColor,
        int TaskbarBevelSize,
        string ThemeProfileName,
        List<string> QuickLaunchOrder,
        bool ShowQuickLaunchToolbar,
        bool ShowAddressToolbar,
        bool AddressToolbarBeforeQuickLaunch,
        bool ShowSpotifyToolbar,
        List<string> AddressHistory,
        List<StartMenuRecentProgramSetting> StartMenuRecentPrograms);

    internal static class TaskbarSettingsService
    {
        private const int AddressHistoryMaxEntries = 40;

        public static TaskbarRuntimeSettings Load()
        {
            var settings = TaskbarSettingsStore.Load();
            var buttonStyle = NormalizeButtonStyle(settings.TaskbarButtonStyle, settings.TaskbarBevelSize, out var bevelSize);

            return new TaskbarRuntimeSettings(
                StartMenuIconSize: Math.Clamp(settings.StartMenuIconSize, 16, 32),
                TaskIconSize: Math.Clamp(settings.TaskIconSize, 16, 32),
                LazyLoadProgramsSubmenu: settings.LazyLoadProgramsSubmenu,
                PlayVolumeFeedbackSound: settings.PlayVolumeFeedbackSound,
                UseClassicVolumePopup: settings.UseClassicVolumePopup,
                StartMenuSubmenuOpenDelayMs: Math.Clamp(settings.StartMenuSubmenuOpenDelayMs, 0, 1500),
                StartMenuRecentProgramsMaxCount: Math.Clamp(settings.StartMenuRecentProgramsMaxCount, 1, 30),
                AutoHideTaskbar: settings.AutoHideTaskbar,
                StartOnWindowsStartup: settings.StartOnWindowsStartup,
                TaskbarButtonStyle: buttonStyle,
                TaskbarFontName: string.IsNullOrWhiteSpace(settings.TaskbarFontName) ? "\u65B0\u7D30\u660E\u9AD4" : settings.TaskbarFontName,
                TaskbarFontSize: Math.Clamp(settings.TaskbarFontSize, 7f, 16f),
                TaskbarFontColor: ParseColorOrDefault(settings.TaskbarFontColor, Color.Black),
                TaskbarLocked: settings.TaskbarLocked,
                TaskbarRows: Math.Clamp(settings.TaskbarRows, 1, 3),
                TaskbarBaseColor: ParseColorOrDefault(settings.TaskbarBaseColor, Color.FromArgb(192, 192, 192)),
                TaskbarLightColor: ParseColorOrDefault(settings.TaskbarLightColor, Color.FromArgb(255, 255, 255)),
                TaskbarDarkColor: ParseColorOrDefault(settings.TaskbarDarkColor, Color.FromArgb(128, 128, 128)),
                TaskbarBevelSize: bevelSize,
                ThemeProfileName: string.IsNullOrWhiteSpace(settings.ThemeProfileName) ? "Custom" : settings.ThemeProfileName,
                QuickLaunchOrder: settings.QuickLaunchOrder?.ToList() ?? new List<string>(),
                ShowQuickLaunchToolbar: settings.ShowQuickLaunchToolbar,
                ShowAddressToolbar: settings.ShowAddressToolbar,
                AddressToolbarBeforeQuickLaunch: settings.AddressToolbarBeforeQuickLaunch,
                ShowSpotifyToolbar: settings.ShowSpotifyToolbar,
                AddressHistory: NormalizeAddressHistory(settings.AddressHistory),
                StartMenuRecentPrograms: NormalizeRecentPrograms(settings.StartMenuRecentPrograms, settings.StartMenuRecentProgramsMaxCount));
        }

        public static void Save(TaskbarRuntimeSettings settings)
        {
            TaskbarSettingsStore.Save(new TaskbarSettings
            {
                StartMenuIconSize = settings.StartMenuIconSize,
                TaskIconSize = settings.TaskIconSize,
                LazyLoadProgramsSubmenu = settings.LazyLoadProgramsSubmenu,
                PlayVolumeFeedbackSound = settings.PlayVolumeFeedbackSound,
                UseClassicVolumePopup = settings.UseClassicVolumePopup,
                StartMenuSubmenuOpenDelayMs = settings.StartMenuSubmenuOpenDelayMs,
                StartMenuRecentProgramsMaxCount = settings.StartMenuRecentProgramsMaxCount,
                AutoHideTaskbar = settings.AutoHideTaskbar,
                StartOnWindowsStartup = settings.StartOnWindowsStartup,
                TaskbarButtonStyle = settings.TaskbarButtonStyle,
                TaskbarFontName = settings.TaskbarFontName,
                TaskbarFontSize = settings.TaskbarFontSize,
                TaskbarFontColor = ColorTranslator.ToHtml(settings.TaskbarFontColor),
                TaskbarLocked = settings.TaskbarLocked,
                TaskbarRows = settings.TaskbarRows,
                TaskbarBaseColor = ColorTranslator.ToHtml(settings.TaskbarBaseColor),
                TaskbarLightColor = ColorTranslator.ToHtml(settings.TaskbarLightColor),
                TaskbarDarkColor = ColorTranslator.ToHtml(settings.TaskbarDarkColor),
                TaskbarBevelSize = settings.TaskbarBevelSize,
                ThemeProfileName = settings.ThemeProfileName,
                QuickLaunchOrder = settings.QuickLaunchOrder?.ToList() ?? new List<string>(),
                ShowQuickLaunchToolbar = settings.ShowQuickLaunchToolbar,
                ShowAddressToolbar = settings.ShowAddressToolbar,
                AddressToolbarBeforeQuickLaunch = settings.AddressToolbarBeforeQuickLaunch,
                ShowSpotifyToolbar = settings.ShowSpotifyToolbar,
                AddressHistory = NormalizeAddressHistory(settings.AddressHistory),
                StartMenuRecentPrograms = NormalizeRecentPrograms(settings.StartMenuRecentPrograms, settings.StartMenuRecentProgramsMaxCount)
            });
        }

        private static string NormalizeButtonStyle(string? value, int configuredBevelSize, out int bevelSize)
        {
            bevelSize = Math.Clamp(configuredBevelSize, 1, 4);

            if (string.Equals(value, "Modern", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(value, "Win98", StringComparison.OrdinalIgnoreCase))
            {
                return string.Equals(value, "Modern", StringComparison.OrdinalIgnoreCase) ? "Modern" : "Win98";
            }

            if (string.Equals(value, "Win98 Thick", StringComparison.OrdinalIgnoreCase))
            {
                bevelSize = Math.Max(bevelSize, 2);
                return "Win98";
            }

            if (string.Equals(value, "Classic", StringComparison.OrdinalIgnoreCase))
            {
                return "Modern";
            }

            if (string.Equals(value, "Flat", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(value, "Borderless", StringComparison.OrdinalIgnoreCase))
            {
                return "Win98";
            }

            return "Modern";
        }

        private static List<string> NormalizeAddressHistory(IEnumerable<string>? history)
        {
            return (history ?? Enumerable.Empty<string>())
                .Select(item => item?.Trim() ?? string.Empty)
                .Where(item => !string.IsNullOrWhiteSpace(item))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Take(AddressHistoryMaxEntries)
                .ToList();
        }

        private static List<StartMenuRecentProgramSetting> NormalizeRecentPrograms(
            IEnumerable<StartMenuRecentProgramSetting>? programs,
            int maxCount)
        {
            return (programs ?? Enumerable.Empty<StartMenuRecentProgramSetting>())
                .Where(item => !string.IsNullOrWhiteSpace(item.LaunchPath))
                .DistinctBy(item => $"{item.IsShellApp}|{item.LaunchPath}", StringComparer.OrdinalIgnoreCase)
                .Take(Math.Clamp(maxCount, 1, 30))
                .ToList();
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
    }
}
