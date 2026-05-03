using ManagedShell.AppBar;
using ManagedShell.WindowsTray;
using TrayNotifyIcon = ManagedShell.WindowsTray.NotifyIcon;

namespace win9xplorer
{
    internal sealed class TrayController
    {
        internal enum TrayFallbackAction
        {
            None,
            OpenWindowsSecurityCenter
        }

        internal sealed record TrayFallbackRule(
            string Name,
            string[] IdentifierContains,
            string[] PathContains,
            string[] TitleContains,
            bool ForceHide = false,
            string? TooltipOverride = null,
            TrayFallbackAction DoubleClickAction = TrayFallbackAction.None);

        private static readonly TrayFallbackRule[] TrayFallbackRules =
        {
            new(
                "Volume",
                new[] { "volume" },
                new[] { "sndvol", "audio", "audiosrv", "systemsettings" },
                new[] { "volume" },
                ForceHide: true,
                TooltipOverride: "Volume"),
            new(
                "WindowsSecurity",
                new[] { "securityhealth" },
                new[] { "securityhealthsystray" },
                new[] { "windows security", "windows 安全", "安全性" },
                TooltipOverride: "Windows Security",
                DoubleClickAction: TrayFallbackAction.OpenWindowsSecurityCenter)
        };

        public bool IsVolumeTrayIcon(TrayNotifyIcon icon)
        {
            if (icon.GUID != Guid.Empty &&
                icon.GUID.ToString().Equals(NotificationArea.VOLUME_GUID, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return TryGetFallbackRule(icon, out var rule) && string.Equals(rule.Name, "Volume", StringComparison.OrdinalIgnoreCase);
        }

        public bool ShouldHideTrayIcon(TrayNotifyIcon icon)
        {
            if (IsVolumeTrayIcon(icon))
            {
                return true;
            }

            return TryGetFallbackRule(icon, out var rule) && rule.ForceHide;
        }

        public bool ShouldOpenWindowsSecurityCenterOnDoubleClick(TrayNotifyIcon icon)
        {
            return TryGetFallbackRule(icon, out var rule) && rule.DoubleClickAction == TrayFallbackAction.OpenWindowsSecurityCenter;
        }

        public bool ShouldSuppressHoverForwarding(TrayNotifyIcon icon)
        {
            return ContainsIgnoreCase(icon.Identifier, "spotify") ||
                   ContainsIgnoreCase(icon.Path, "spotify") ||
                   ContainsIgnoreCase(icon.Title, "spotify");
        }

        public bool HasStableIdentity(TrayNotifyIcon icon)
        {
            return icon.GUID != Guid.Empty ||
                   !string.IsNullOrWhiteSpace(icon.Identifier) ||
                   !string.IsNullOrWhiteSpace(icon.Path) ||
                   !string.IsNullOrWhiteSpace(icon.Title);
        }

        public string GetHoverMessage(TrayNotifyIcon icon)
        {
            if (TryGetFallbackRule(icon, out var rule) && !string.IsNullOrWhiteSpace(rule.TooltipOverride))
            {
                return rule.TooltipOverride;
            }

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

        private static bool TryGetFallbackRule(TrayNotifyIcon icon, out TrayFallbackRule rule)
        {
            foreach (var candidate in TrayFallbackRules)
            {
                if (MatchesRule(icon, candidate))
                {
                    rule = candidate;
                    return true;
                }
            }

            rule = null!;
            return false;
        }

        private static bool MatchesRule(TrayNotifyIcon icon, TrayFallbackRule rule)
        {
            return rule.IdentifierContains.Any(token => ContainsIgnoreCase(icon.Identifier, token)) ||
                   rule.PathContains.Any(token => ContainsIgnoreCase(icon.Path, token)) ||
                   rule.TitleContains.Any(token => ContainsIgnoreCase(icon.Title, token));
        }

        private static bool ContainsIgnoreCase(string? source, string token)
        {
            return !string.IsNullOrWhiteSpace(source) && source.Contains(token, StringComparison.OrdinalIgnoreCase);
        }
    }
}
