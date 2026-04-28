using System.ComponentModel;

namespace win9xplorer
{
    internal sealed class QuickLaunchRefreshMonitor : IDisposable
    {
        private readonly ISynchronizeInvoke synchronizer;
        private readonly System.Windows.Forms.Timer debounceTimer;
        private FileSystemWatcher? watcher;

        public event EventHandler? RefreshRequested;

        public QuickLaunchRefreshMonitor(ISynchronizeInvoke synchronizer, int debounceMilliseconds = 300)
        {
            this.synchronizer = synchronizer;
            debounceTimer = new System.Windows.Forms.Timer { Interval = debounceMilliseconds };
            debounceTimer.Tick += (_, _) =>
            {
                debounceTimer.Stop();
                RefreshRequested?.Invoke(this, EventArgs.Empty);
            };
        }

        public void Start()
        {
            var quickLaunchFolder = GetQuickLaunchFolderPath();
            Directory.CreateDirectory(quickLaunchFolder);

            watcher = new FileSystemWatcher(quickLaunchFolder)
            {
                IncludeSubdirectories = false,
                NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite | NotifyFilters.CreationTime,
                EnableRaisingEvents = true
            };

            watcher.Changed += Watcher_Changed;
            watcher.Created += Watcher_Changed;
            watcher.Deleted += Watcher_Changed;
            watcher.Renamed += Watcher_Renamed;
        }

        private void Watcher_Changed(object sender, FileSystemEventArgs e)
        {
            ScheduleRefresh();
        }

        private void Watcher_Renamed(object sender, RenamedEventArgs e)
        {
            ScheduleRefresh();
        }

        private void ScheduleRefresh()
        {
            if (synchronizer.InvokeRequired)
            {
                synchronizer.BeginInvoke(new Action(ScheduleRefresh), Array.Empty<object>());
                return;
            }

            debounceTimer.Stop();
            debounceTimer.Start();
        }

        public void Dispose()
        {
            if (watcher != null)
            {
                watcher.Changed -= Watcher_Changed;
                watcher.Created -= Watcher_Changed;
                watcher.Deleted -= Watcher_Changed;
                watcher.Renamed -= Watcher_Renamed;
                watcher.Dispose();
                watcher = null;
            }

            debounceTimer.Dispose();
        }

        private static string GetQuickLaunchFolderPath()
        {
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "Microsoft",
                "Internet Explorer",
                "Quick Launch");
        }
    }
}
