using System.Text.Json;

namespace win9xplorer
{
    internal sealed class TaskbarSettings
    {
        public int StartMenuIconSize { get; set; } = 20;

        public int TaskIconSize { get; set; } = 20;

        public bool LazyLoadProgramsSubmenu { get; set; } = true;

        public bool PlayVolumeFeedbackSound { get; set; } = true;

        public int StartMenuSubmenuOpenDelayMs { get; set; } = 200;

        public bool AutoHideTaskbar { get; set; } = false;

        public string TaskbarButtonStyle { get; set; } = "Classic";

        public string TaskbarFontName { get; set; } = "MS Sans Serif";

        public float TaskbarFontSize { get; set; } = 8.25f;

        public bool TaskbarLocked { get; set; } = false;

        public int TaskbarRows { get; set; } = 1;

        public string TaskbarBaseColor { get; set; } = "#C0C0C0";

        public string TaskbarLightColor { get; set; } = "#FFFFFF";

        public string TaskbarDarkColor { get; set; } = "#808080";

        public int TaskbarBevelSize { get; set; } = 1;
    }

    internal static class TaskbarSettingsStore
    {
        private static readonly string SettingsPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "win9xplorer",
            "taskbar.settings.json");

        public static TaskbarSettings Load()
        {
            try
            {
                if (!File.Exists(SettingsPath))
                {
                    return new TaskbarSettings();
                }

                var json = File.ReadAllText(SettingsPath);
                var settings = JsonSerializer.Deserialize<TaskbarSettings>(json);
                return settings ?? new TaskbarSettings();
            }
            catch
            {
                return new TaskbarSettings();
            }
        }

        public static void Save(TaskbarSettings settings)
        {
            try
            {
                var folder = Path.GetDirectoryName(SettingsPath);
                if (!string.IsNullOrWhiteSpace(folder))
                {
                    Directory.CreateDirectory(folder);
                }

                var json = JsonSerializer.Serialize(settings, new JsonSerializerOptions
                {
                    WriteIndented = true
                });
                File.WriteAllText(SettingsPath, json);
            }
            catch
            {
            }
        }
    }
}
