using System.Diagnostics;
using System.Runtime.InteropServices;

namespace win9xplorer
{
    internal sealed class SpotifyController
    {
        private const int KeyEventKeyUp = 0x0002;
        private const byte VkMediaNextTrack = 0xB0;
        private const byte VkMediaPrevTrack = 0xB1;
        private const byte VkMediaPlayPause = 0xB3;

        [DllImport("user32.dll", SetLastError = true)]
        private static extern void keybd_event(byte bVk, byte bScan, int dwFlags, IntPtr dwExtraInfo);

        internal enum SpotifyCommand
        {
            PlayPause,
            PreviousTrack,
            NextTrack
        }

        public bool IsSpotifyRunning()
        {
            return Process.GetProcessesByName("Spotify").Length > 0;
        }

        public string GetNowPlayingTitle()
        {
            try
            {
                return Process.GetProcessesByName("Spotify")
                    .Select(process =>
                    {
                        using (process)
                        {
                            return process.MainWindowTitle?.Trim() ?? string.Empty;
                        }
                    })
                    .Where(title => !string.IsNullOrWhiteSpace(title))
                    .FirstOrDefault(title =>
                        !title.Equals("Spotify", StringComparison.OrdinalIgnoreCase) &&
                        !title.Equals("Spotify Premium", StringComparison.OrdinalIgnoreCase) &&
                        !title.Equals("Spotify Free", StringComparison.OrdinalIgnoreCase))
                    ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        public bool TryExecute(SpotifyCommand command)
        {
            if (!IsSpotifyRunning())
            {
                return false;
            }

            var keyCode = command switch
            {
                SpotifyCommand.PlayPause => VkMediaPlayPause,
                SpotifyCommand.PreviousTrack => VkMediaPrevTrack,
                SpotifyCommand.NextTrack => VkMediaNextTrack,
                _ => VkMediaPlayPause
            };

            keybd_event(keyCode, 0, 0, IntPtr.Zero);
            keybd_event(keyCode, 0, KeyEventKeyUp, IntPtr.Zero);
            return true;
        }
    }
}
