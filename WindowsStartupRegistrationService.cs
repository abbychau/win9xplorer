using System.Diagnostics;
using Microsoft.Win32;

namespace win9xplorer
{
    internal static class WindowsStartupRegistrationService
    {
        private const string RunPath = @"Software\Microsoft\Windows\CurrentVersion\Run";
        private const string RunName = "win9xplorer";

        public static void Apply(bool enabled, string executablePath)
        {
            try
            {
                using var key = Registry.CurrentUser.CreateSubKey(RunPath);
                if (key == null)
                {
                    return;
                }

                if (enabled)
                {
                    key.SetValue(RunName, $"\"{executablePath}\"");
                }
                else
                {
                    key.DeleteValue(RunName, false);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to update startup registration: {ex.Message}");
            }
        }
    }
}
