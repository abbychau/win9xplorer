namespace win9xplorer
{
    internal sealed class QuickLaunchController
    {
        public sealed record QuickLaunchEntry(string FullPath, bool IsSeparator);

        public List<QuickLaunchEntry> GetOrderedEntries(string folderPath, IReadOnlyCollection<string>? preferredOrder)
        {
            var launchFiles = new[] { "*.lnk", "*.url", "*.scf", "*.separator" }
                .SelectMany(pattern => Directory.GetFiles(folderPath, pattern))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            var priorityMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            if (preferredOrder != null)
            {
                var index = 0;
                foreach (var name in preferredOrder)
                {
                    if (string.IsNullOrWhiteSpace(name) || priorityMap.ContainsKey(name))
                    {
                        continue;
                    }

                    priorityMap[name] = index++;
                }
            }

            return launchFiles
                .Select(path => new QuickLaunchEntry(path, Path.GetExtension(path).Equals(".separator", StringComparison.OrdinalIgnoreCase)))
                .OrderBy(entry => priorityMap.TryGetValue(Path.GetFileName(entry.FullPath), out var position) ? position : int.MaxValue)
                .ThenBy(entry => Path.GetFileNameWithoutExtension(entry.FullPath), StringComparer.CurrentCultureIgnoreCase)
                .ToList();
        }

        public List<string> GetOrderFromPanel(FlowLayoutPanel panel)
        {
            return panel.Controls
                .OfType<Control>()
                .Where(control => control.Tag is string path && !string.IsNullOrWhiteSpace(path))
                .OrderBy(control => control.Left)
                .ThenBy(control => control.Top)
                .Select(control => Path.GetFileName((string)control.Tag!))
                .ToList();
        }

        public int GetDropInsertPosition(IReadOnlyList<Control> controls, Point dropPoint)
        {
            for (var i = 0; i < controls.Count; i++)
            {
                var midpoint = controls[i].Left + (controls[i].Width / 2);
                if (dropPoint.X < midpoint)
                {
                    return i;
                }
            }

            return controls.Count;
        }
    }
}
