using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Security.Principal;

namespace win9xplorer
{
    internal sealed class TaskManagerForm : Form
    {
        private const int RefreshIntervalMs = 1500;
        private const int ProcessQueryLimitedInformation = 0x1000;
        private const int TokenQuery = 0x0008;

        private readonly TabControl tabs;
        private readonly TextBox filterTextBox;
        private readonly Label summaryLabel;
        private readonly ListView processList;
        private readonly ContextMenuStrip processContextMenu;
        private readonly CheckBox showAllUsersCheckBox;
        private readonly Button endProcessButton;
        private readonly StatusStrip statusStrip;
        private readonly ToolStripStatusLabel processCountStatus;
        private readonly ToolStripStatusLabel cpuStatus;
        private readonly ToolStripStatusLabel memoryStatus;
        private readonly System.Windows.Forms.Timer refreshTimer;
        private readonly Dictionary<int, ProcessSample> previousSamples = new();
        private readonly string currentUserName;

        private List<ProcessRow> allRows = new();
        private DateTime previousRefreshUtc = DateTime.UtcNow;
        private int sortColumn = 2;
        private SortOrder sortOrder = SortOrder.Descending;
        private bool refreshPaused;

        public TaskManagerForm()
        {
            Text = "win9xplorer Task Manager";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.Sizable;
            MinimizeBox = true;
            MaximizeBox = true;
            ShowInTaskbar = true;
            ClientSize = new Size(820, 500);
            MinimumSize = new Size(620, 360);
            Icon = LoadTaskManagerIcon();
            currentUserName = GetCurrentShortUserName();

            MainMenuStrip = BuildMenuStrip();
            Controls.Add(MainMenuStrip);

            tabs = new TabControl
            {
                Left = 8,
                Top = MainMenuStrip.Bottom + 4,
                Width = ClientSize.Width - 16,
                Height = ClientSize.Height - MainMenuStrip.Height - 58,
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                Font = Font
            };
            Controls.Add(tabs);

            var processTab = new TabPage("Processes") { Padding = new Padding(8) };
            tabs.TabPages.Add(processTab);

            var tabLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2
            };
            tabLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            tabLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            processTab.Controls.Add(tabLayout);

            var filterPanel = new Panel { Dock = DockStyle.Fill };
            tabLayout.Controls.Add(filterPanel, 0, 0);

            var filterLabel = new Label
            {
                AutoSize = true,
                Left = 0,
                Top = 6,
                Text = "Filter:"
            };
            filterPanel.Controls.Add(filterLabel);

            filterTextBox = new TextBox
            {
                Left = filterLabel.Right + 8,
                Top = 2,
                Width = 210,
                Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Bottom
            };
            filterTextBox.TextChanged += (_, _) => ApplyView();
            filterPanel.Controls.Add(filterTextBox);

            summaryLabel = new Label
            {
                Left = filterTextBox.Right + 12,
                Top = 6,
                Width = 420,
                Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right,
                AutoEllipsis = true
            };
            filterPanel.Controls.Add(summaryLabel);
            filterPanel.Resize += (_, _) =>
            {
                summaryLabel.Left = filterTextBox.Right + 12;
                summaryLabel.Width = Math.Max(120, filterPanel.ClientSize.Width - summaryLabel.Left);
            };

            processList = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                HideSelection = false,
                MultiSelect = false,
                BorderStyle = BorderStyle.Fixed3D,
                Font = Font
            };
            processList.Columns.Add("Image Name", 170, HorizontalAlignment.Left);
            processList.Columns.Add("PID", 64, HorizontalAlignment.Right);
            processList.Columns.Add("CPU", 50, HorizontalAlignment.Right);
            processList.Columns.Add("Memory", 86, HorizontalAlignment.Right);
            processList.Columns.Add("User Name", 100, HorizontalAlignment.Left);
            processList.Columns.Add("Threads", 64, HorizontalAlignment.Right);
            processList.Columns.Add("Handles", 72, HorizontalAlignment.Right);
            processList.Columns.Add("Start Time", 130, HorizontalAlignment.Left);
            processList.Columns.Add("Description", 180, HorizontalAlignment.Left);
            processList.Columns.Add("Path", 260, HorizontalAlignment.Left);
            processList.ColumnClick += ProcessList_ColumnClick;
            processList.SelectedIndexChanged += (_, _) => UpdateSelectionState();
            processList.DoubleClick += (_, _) => OpenSelectedProcessLocation();
            tabLayout.Controls.Add(processList, 0, 1);

            processContextMenu = BuildProcessContextMenu();
            processList.ContextMenuStrip = processContextMenu;

            showAllUsersCheckBox = new CheckBox
            {
                Left = 8,
                Top = ClientSize.Height - 50,
                Width = 220,
                Height = 24,
                Anchor = AnchorStyles.Left | AnchorStyles.Bottom,
                Text = "Show processes from all users",
                Checked = true,
                Font = Font
            };
            showAllUsersCheckBox.CheckedChanged += (_, _) => ApplyView();
            Controls.Add(showAllUsersCheckBox);

            endProcessButton = new Button
            {
                Text = "End Process",
                Width = 104,
                Height = 27,
                Left = ClientSize.Width - 120,
                Top = ClientSize.Height - 52,
                Anchor = AnchorStyles.Right | AnchorStyles.Bottom,
                Enabled = false,
                Font = Font
            };
            endProcessButton.Click += (_, _) => EndSelectedProcess();
            Controls.Add(endProcessButton);

            statusStrip = new StatusStrip
            {
                SizingGrip = true,
                Font = Font
            };
            processCountStatus = new ToolStripStatusLabel { BorderSides = ToolStripStatusLabelBorderSides.Right };
            cpuStatus = new ToolStripStatusLabel { BorderSides = ToolStripStatusLabelBorderSides.Right };
            memoryStatus = new ToolStripStatusLabel();
            statusStrip.Items.Add(processCountStatus);
            statusStrip.Items.Add(cpuStatus);
            statusStrip.Items.Add(memoryStatus);
            Controls.Add(statusStrip);

            refreshTimer = new System.Windows.Forms.Timer { Interval = RefreshIntervalMs };
            refreshTimer.Tick += (_, _) =>
            {
                if (!refreshPaused)
                {
                    RefreshProcesses();
                }
            };

            Load += (_, _) =>
            {
                RefreshProcesses();
                refreshTimer.Start();
            };

            FormClosed += (_, _) =>
            {
                refreshTimer.Stop();
                refreshTimer.Dispose();
                processContextMenu.Dispose();
            };
        }

        private MenuStrip BuildMenuStrip()
        {
            var menu = new MenuStrip
            {
                Font = Font,
                GripStyle = ToolStripGripStyle.Hidden
            };

            var fileMenu = new ToolStripMenuItem("File");
            fileMenu.DropDownItems.Add("Run new task...", null, (_, _) => RunNewTask());
            fileMenu.DropDownItems.Add(new ToolStripSeparator());
            fileMenu.DropDownItems.Add("Exit", null, (_, _) => Close());

            var optionsMenu = new ToolStripMenuItem("Options");
            var pauseItem = new ToolStripMenuItem("Pause Refresh") { CheckOnClick = true };
            pauseItem.CheckedChanged += (_, _) =>
            {
                refreshPaused = pauseItem.Checked;
                UpdateStatus();
            };
            optionsMenu.DropDownItems.Add(pauseItem);

            var viewMenu = new ToolStripMenuItem("View");
            viewMenu.DropDownItems.Add("Refresh Now\tF5", null, (_, _) => RefreshProcesses());
            viewMenu.DropDownItems.Add("Clear Filter\tEsc", null, (_, _) => ClearFilter());

            var helpMenu = new ToolStripMenuItem("Help");
            helpMenu.DropDownItems.Add("About", null, (_, _) => MessageBox.Show(
                "win9xplorer Task Manager shows a focused process view with sortable power-user columns.",
                "win9xplorer Task Manager",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information));

            menu.Items.Add(fileMenu);
            menu.Items.Add(optionsMenu);
            menu.Items.Add(viewMenu);
            menu.Items.Add(helpMenu);
            return menu;
        }

        private ContextMenuStrip BuildProcessContextMenu()
        {
            var menu = new ContextMenuStrip { Font = Font };
            menu.Items.Add("End Process", null, (_, _) => EndSelectedProcess());
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("Open File Location", null, (_, _) => OpenSelectedProcessLocation());
            menu.Items.Add("Copy Details", null, (_, _) => CopySelectedProcessDetails());
            menu.Items.Add("Copy Path", null, (_, _) => CopySelectedProcessPath());
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("Refresh", null, (_, _) => RefreshProcesses());
            menu.Opening += (_, e) =>
            {
                var hasSelection = GetSelectedRow() != null;
                foreach (ToolStripItem item in menu.Items)
                {
                    if (item is ToolStripSeparator)
                    {
                        continue;
                    }

                    item.Enabled = hasSelection || string.Equals(item.Text, "Refresh", StringComparison.OrdinalIgnoreCase);
                }
            };
            return menu;
        }

        private void RefreshProcesses()
        {
            var now = DateTime.UtcNow;
            var elapsedSeconds = Math.Max(0.1, (now - previousRefreshUtc).TotalSeconds);
            var processorCount = Math.Max(1, Environment.ProcessorCount);
            var selectedPid = GetSelectedProcessId();
            var currentSamples = new Dictionary<int, ProcessSample>();
            var rows = new List<ProcessRow>();

            foreach (var process in Process.GetProcesses().OrderBy(p => p.ProcessName, StringComparer.OrdinalIgnoreCase))
            {
                using (process)
                {
                    try
                    {
                        var pid = process.Id;
                        var processPath = GetProcessPath(process);
                        var totalProcessorTime = GetTotalProcessorTime(process);
                        var workingSet = GetWorkingSet(process);
                        var sample = new ProcessSample(totalProcessorTime, workingSet);
                        currentSamples[pid] = sample;

                        var cpu = 0;
                        if (previousSamples.TryGetValue(pid, out var previous))
                        {
                            var cpuPercent = (totalProcessorTime - previous.TotalProcessorTime).TotalSeconds / elapsedSeconds / processorCount * 100;
                            cpu = Math.Clamp((int)Math.Round(cpuPercent), 0, 100);
                        }

                        rows.Add(new ProcessRow(
                            pid,
                            GetImageName(process, processPath),
                            GetProcessUserName(pid),
                            cpu,
                            workingSet,
                            GetThreadCount(process),
                            GetHandleCount(process),
                            GetStartTime(process),
                            GetDescription(processPath),
                            processPath));
                    }
                    catch
                    {
                        // Processes can exit or deny access while the list is being built.
                    }
                }
            }

            allRows = rows;
            ApplyView(selectedPid);

            previousSamples.Clear();
            foreach (var pair in currentSamples)
            {
                previousSamples[pair.Key] = pair.Value;
            }
            previousRefreshUtc = now;
        }

        private void ApplyView(int? preferredSelectedPid = null)
        {
            var filter = filterTextBox.Text.Trim();
            var selectedPid = preferredSelectedPid ?? GetSelectedProcessId();
            var visibleRows = allRows.Where(row => ShouldShowRow(row, filter)).ToList();
            visibleRows.Sort(CompareRows);

            processList.BeginUpdate();
            processList.Items.Clear();
            foreach (var row in visibleRows)
            {
                var item = new ListViewItem(row.ImageName) { Tag = row };
                item.SubItems.Add(row.ProcessId.ToString());
                item.SubItems.Add(row.Cpu.ToString("00"));
                item.SubItems.Add(FormatBytes(row.WorkingSetBytes));
                item.SubItems.Add(row.UserName);
                item.SubItems.Add(row.ThreadCount.ToString());
                item.SubItems.Add(row.HandleCount?.ToString() ?? string.Empty);
                item.SubItems.Add(row.StartTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? string.Empty);
                item.SubItems.Add(row.Description);
                item.SubItems.Add(row.Path);
                processList.Items.Add(item);

                if (row.ProcessId == selectedPid)
                {
                    item.Selected = true;
                    item.Focused = true;
                }
            }
            processList.EndUpdate();

            if (processList.SelectedItems.Count > 0)
            {
                processList.SelectedItems[0].EnsureVisible();
            }

            UpdateSelectionState();
            UpdateStatus(visibleRows);
        }

        private bool ShouldShowRow(ProcessRow row, string filter)
        {
            if (!showAllUsersCheckBox.Checked &&
                !string.Equals(row.UserName, currentUserName, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            if (string.IsNullOrWhiteSpace(filter))
            {
                return true;
            }

            return row.ImageName.Contains(filter, StringComparison.OrdinalIgnoreCase) ||
                   row.ProcessId.ToString().Contains(filter, StringComparison.OrdinalIgnoreCase) ||
                   row.UserName.Contains(filter, StringComparison.OrdinalIgnoreCase) ||
                   row.Description.Contains(filter, StringComparison.OrdinalIgnoreCase) ||
                   row.Path.Contains(filter, StringComparison.OrdinalIgnoreCase);
        }

        private void UpdateSelectionState()
        {
            var row = GetSelectedRow();
            endProcessButton.Enabled = row != null && row.ProcessId != Environment.ProcessId;
            summaryLabel.Text = row == null
                ? "Select a process to inspect its executable path."
                : $"{row.ImageName}  PID {row.ProcessId}  {FormatBytes(row.WorkingSetBytes)}  {row.Path}";
        }

        private void UpdateStatus(IReadOnlyCollection<ProcessRow>? visibleRows = null)
        {
            var rows = visibleRows ?? allRows.Where(row => ShouldShowRow(row, filterTextBox.Text.Trim())).ToList();
            var totalCpu = rows.Sum(row => row.Cpu);
            var pausedText = refreshPaused ? " (paused)" : string.Empty;
            processCountStatus.Text = $"Processes: {rows.Count}/{allRows.Count}{pausedText}";
            cpuStatus.Text = $"CPU Usage: {Math.Clamp(totalCpu, 0, 100)}%";
            memoryStatus.Text = $"Physical Memory: {GetPhysicalMemoryLoad()}%";
        }

        private void ProcessList_ColumnClick(object? sender, ColumnClickEventArgs e)
        {
            if (sortColumn == e.Column)
            {
                sortOrder = sortOrder == SortOrder.Ascending ? SortOrder.Descending : SortOrder.Ascending;
            }
            else
            {
                sortColumn = e.Column;
                sortOrder = e.Column is 2 or 3 or 5 or 6 ? SortOrder.Descending : SortOrder.Ascending;
            }

            ApplyView();
        }

        private int CompareRows(ProcessRow left, ProcessRow right)
        {
            var result = sortColumn switch
            {
                0 => string.Compare(left.ImageName, right.ImageName, StringComparison.OrdinalIgnoreCase),
                1 => left.ProcessId.CompareTo(right.ProcessId),
                2 => left.Cpu.CompareTo(right.Cpu),
                3 => left.WorkingSetBytes.CompareTo(right.WorkingSetBytes),
                4 => string.Compare(left.UserName, right.UserName, StringComparison.OrdinalIgnoreCase),
                5 => left.ThreadCount.CompareTo(right.ThreadCount),
                6 => Nullable.Compare(left.HandleCount, right.HandleCount),
                7 => Nullable.Compare(left.StartTime, right.StartTime),
                8 => string.Compare(left.Description, right.Description, StringComparison.OrdinalIgnoreCase),
                9 => string.Compare(left.Path, right.Path, StringComparison.OrdinalIgnoreCase),
                _ => 0
            };

            if (result == 0)
            {
                result = left.ProcessId.CompareTo(right.ProcessId);
            }

            return sortOrder == SortOrder.Descending ? -result : result;
        }

        private int? GetSelectedProcessId()
        {
            return GetSelectedRow()?.ProcessId;
        }

        private ProcessRow? GetSelectedRow()
        {
            if (processList.SelectedItems.Count == 0 || processList.SelectedItems[0].Tag is not ProcessRow row)
            {
                return null;
            }

            return row;
        }

        private void EndSelectedProcess()
        {
            var row = GetSelectedRow();
            if (row == null)
            {
                return;
            }

            if (row.ProcessId == Environment.ProcessId)
            {
                MessageBox.Show("Refusing to end the running win9xplorer process from inside its own Task Manager.", "Task Manager", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var result = MessageBox.Show(
                $"Ending this process can cause data loss or system instability.{Environment.NewLine}{Environment.NewLine}{row.ImageName} (PID {row.ProcessId}){Environment.NewLine}{row.Path}{Environment.NewLine}{Environment.NewLine}End this process?",
                "Task Manager Warning",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button2);
            if (result != DialogResult.Yes)
            {
                return;
            }

            try
            {
                using var process = Process.GetProcessById(row.ProcessId);
                process.Kill();
                process.WaitForExit(500);
                RefreshProcesses();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Unable to end process.{Environment.NewLine}{ex.Message}", "Task Manager", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OpenSelectedProcessLocation()
        {
            var row = GetSelectedRow();
            if (row == null || string.IsNullOrWhiteSpace(row.Path) || !File.Exists(row.Path))
            {
                System.Media.SystemSounds.Beep.Play();
                return;
            }

            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "explorer.exe",
                    Arguments = $"/select,\"{row.Path}\"",
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Unable to open file location.{Environment.NewLine}{ex.Message}", "Task Manager", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CopySelectedProcessDetails()
        {
            var row = GetSelectedRow();
            if (row == null)
            {
                return;
            }

            Clipboard.SetText(string.Join(Environment.NewLine,
                $"Image: {row.ImageName}",
                $"PID: {row.ProcessId}",
                $"User: {row.UserName}",
                $"CPU: {row.Cpu:00}",
                $"Memory: {FormatBytes(row.WorkingSetBytes)}",
                $"Threads: {row.ThreadCount}",
                $"Handles: {row.HandleCount?.ToString() ?? string.Empty}",
                $"Start time: {row.StartTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? string.Empty}",
                $"Description: {row.Description}",
                $"Path: {row.Path}"));
        }

        private void CopySelectedProcessPath()
        {
            var row = GetSelectedRow();
            if (row == null || string.IsNullOrWhiteSpace(row.Path))
            {
                return;
            }

            Clipboard.SetText(row.Path);
        }

        private void RunNewTask()
        {
            using var dialog = new Form
            {
                Text = "Create New Task",
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MinimizeBox = false,
                MaximizeBox = false,
                ShowInTaskbar = false,
                ClientSize = new Size(420, 118),
                Font = Font
            };

            var label = new Label
            {
                Left = 12,
                Top = 14,
                Width = 390,
                Text = "Type the name of a program, folder, document, or Internet resource:"
            };
            var input = new TextBox
            {
                Left = 12,
                Top = 40,
                Width = 396
            };
            var okButton = new Button
            {
                Text = "OK",
                DialogResult = DialogResult.OK,
                Left = 252,
                Top = 78,
                Width = 75
            };
            var cancelButton = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Left = 333,
                Top = 78,
                Width = 75
            };

            dialog.Controls.Add(label);
            dialog.Controls.Add(input);
            dialog.Controls.Add(okButton);
            dialog.Controls.Add(cancelButton);
            dialog.AcceptButton = okButton;
            dialog.CancelButton = cancelButton;

            if (dialog.ShowDialog(this) != DialogResult.OK || string.IsNullOrWhiteSpace(input.Text))
            {
                return;
            }

            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = input.Text.Trim(),
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Unable to run task.{Environment.NewLine}{ex.Message}", "Task Manager", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ClearFilter()
        {
            if (filterTextBox.TextLength == 0)
            {
                return;
            }

            filterTextBox.Clear();
            filterTextBox.Focus();
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            switch (keyData)
            {
                case Keys.F5:
                    RefreshProcesses();
                    return true;
                case Keys.Delete:
                    EndSelectedProcess();
                    return true;
                case Keys.Control | Keys.C:
                    CopySelectedProcessDetails();
                    return true;
                case Keys.Control | Keys.F:
                    filterTextBox.Focus();
                    filterTextBox.SelectAll();
                    return true;
                case Keys.Escape:
                    ClearFilter();
                    return true;
            }

            return base.ProcessCmdKey(ref msg, keyData);
        }

        private static string GetProcessPath(Process process)
        {
            try
            {
                return process.MainModule?.FileName ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string GetImageName(Process process, string processPath)
        {
            if (!string.IsNullOrWhiteSpace(processPath))
            {
                return Path.GetFileName(processPath);
            }

            return process.ProcessName.EndsWith(".exe", StringComparison.OrdinalIgnoreCase)
                ? process.ProcessName
                : process.ProcessName + ".exe";
        }

        private static string GetDescription(string processPath)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(processPath))
                {
                    var description = FileVersionInfo.GetVersionInfo(processPath).FileDescription;
                    if (!string.IsNullOrWhiteSpace(description))
                    {
                        return description;
                    }
                }
            }
            catch
            {
            }

            return string.Empty;
        }

        private static TimeSpan GetTotalProcessorTime(Process process)
        {
            try
            {
                return process.TotalProcessorTime;
            }
            catch
            {
                return TimeSpan.Zero;
            }
        }

        private static long GetWorkingSet(Process process)
        {
            try
            {
                return process.WorkingSet64;
            }
            catch
            {
                return 0;
            }
        }

        private static int GetThreadCount(Process process)
        {
            try
            {
                return process.Threads.Count;
            }
            catch
            {
                return 0;
            }
        }

        private static int? GetHandleCount(Process process)
        {
            try
            {
                return process.HandleCount;
            }
            catch
            {
                return null;
            }
        }

        private static DateTime? GetStartTime(Process process)
        {
            try
            {
                return process.StartTime;
            }
            catch
            {
                return null;
            }
        }

        private static int GetPhysicalMemoryLoad()
        {
            var status = new MemoryStatusEx();
            status.Length = (uint)Marshal.SizeOf<MemoryStatusEx>();
            return GlobalMemoryStatusEx(ref status) ? (int)status.MemoryLoad : 0;
        }

        private static string GetProcessUserName(int processId)
        {
            var processHandle = IntPtr.Zero;
            var tokenHandle = IntPtr.Zero;

            try
            {
                processHandle = OpenProcess(ProcessQueryLimitedInformation, false, processId);
                if (processHandle == IntPtr.Zero || !OpenProcessToken(processHandle, TokenQuery, out tokenHandle))
                {
                    return string.Empty;
                }

                using var identity = new WindowsIdentity(tokenHandle);
                return ToShortUserName(identity.Name);
            }
            catch
            {
                return string.Empty;
            }
            finally
            {
                if (tokenHandle != IntPtr.Zero)
                {
                    CloseHandle(tokenHandle);
                }

                if (processHandle != IntPtr.Zero)
                {
                    CloseHandle(processHandle);
                }
            }
        }

        private static string GetCurrentShortUserName()
        {
            try
            {
                return ToShortUserName(WindowsIdentity.GetCurrent().Name);
            }
            catch
            {
                return Environment.UserName;
            }
        }

        private static string ToShortUserName(string name)
        {
            var slash = name.LastIndexOf('\\');
            return slash >= 0 && slash < name.Length - 1 ? name[(slash + 1)..] : name;
        }

        private static string FormatBytes(long bytes)
        {
            return $"{Math.Max(0, bytes / 1024):N0} K";
        }

        private static Icon LoadTaskManagerIcon()
        {
            var candidates = new[]
            {
                Path.Combine(AppContext.BaseDirectory, "Windows 9x Icons", "Task Manager.ico"),
                Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "Windows 9x Icons", "Task Manager.ico")),
                Path.Combine(Environment.CurrentDirectory, "Windows 9x Icons", "Task Manager.ico")
            };

            foreach (var path in candidates)
            {
                try
                {
                    if (File.Exists(path))
                    {
                        return new Icon(path);
                    }
                }
                catch
                {
                }
            }

            return SystemIcons.Application;
        }

        [DllImport("kernel32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GlobalMemoryStatusEx(ref MemoryStatusEx lpBuffer);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr OpenProcess(uint processAccess, bool inheritHandle, int processId);

        [DllImport("advapi32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool OpenProcessToken(IntPtr processHandle, uint desiredAccess, out IntPtr tokenHandle);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool CloseHandle(IntPtr handle);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private struct MemoryStatusEx
        {
            public uint Length;
            public uint MemoryLoad;
            public ulong TotalPhys;
            public ulong AvailPhys;
            public ulong TotalPageFile;
            public ulong AvailPageFile;
            public ulong TotalVirtual;
            public ulong AvailVirtual;
            public ulong AvailExtendedVirtual;
        }

        private sealed record ProcessSample(TimeSpan TotalProcessorTime, long WorkingSetBytes);

        private sealed record ProcessRow(
            int ProcessId,
            string ImageName,
            string UserName,
            int Cpu,
            long WorkingSetBytes,
            int ThreadCount,
            int? HandleCount,
            DateTime? StartTime,
            string Description,
            string Path);
    }
}
