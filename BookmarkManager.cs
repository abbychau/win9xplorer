using System.Diagnostics;

namespace win9xplorer
{
    /// <summary>
    /// Manages bookmarks/favorites for the file explorer
    /// </summary>
    internal class BookmarkManager
    {
        private readonly List<Bookmark> bookmarks = new();
        private readonly RegistrySettingsManager registryManager;
        private Action? refreshTreeViewCallback;

        public BookmarkManager()
        {
            registryManager = new RegistrySettingsManager();
            LoadBookmarks();
        }

        public void SetTreeViewRefreshCallback(Action callback)
        {
            refreshTreeViewCallback = callback;
        }

        public class Bookmark
        {
            public string Name { get; set; } = "";
            public string Path { get; set; } = "";
            public DateTime DateAdded { get; set; }

            public Bookmark() { }

            public Bookmark(string name, string path)
            {
                Name = name;
                Path = path;
                DateAdded = DateTime.Now;
            }

            public Bookmark(string name, string path, DateTime dateAdded)
            {
                Name = name;
                Path = path;
                DateAdded = dateAdded;
            }
        }

        /// <summary>
        /// Add a new bookmark
        /// </summary>
        public void AddBookmark(string path, string? customName = null)
        {
            // Don't add duplicate bookmarks
            if (bookmarks.Any(b => b.Path.Equals(path, StringComparison.OrdinalIgnoreCase)))
            {
                return;
            }

            string name = customName ?? GetFriendlyName(path);
            var bookmark = new Bookmark(name, path);
            bookmarks.Add(bookmark);
            SaveBookmarks();
            refreshTreeViewCallback?.Invoke();

            Debug.WriteLine($"Added bookmark: {name} -> {path}");
        }

        /// <summary>
        /// Remove a bookmark by path
        /// </summary>
        public void RemoveBookmark(string path)
        {
            var bookmark = bookmarks.FirstOrDefault(b => b.Path.Equals(path, StringComparison.OrdinalIgnoreCase));
            if (bookmark != null)
            {
                bookmarks.Remove(bookmark);
                SaveBookmarks();
                refreshTreeViewCallback?.Invoke();
                Debug.WriteLine($"Removed bookmark: {bookmark.Name}");
            }
        }

        /// <summary>
        /// Check if a path is bookmarked
        /// </summary>
        public bool IsBookmarked(string path)
        {
            return bookmarks.Any(b => b.Path.Equals(path, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Get all bookmarks
        /// </summary>
        public List<Bookmark> GetBookmarks()
        {
            return bookmarks.OrderBy(b => b.Name).ToList();
        }

        /// <summary>
        /// Get friendly name for a path
        /// </summary>
        private string GetFriendlyName(string path)
        {
            if (string.IsNullOrEmpty(path) || path == "My Computer")
                return "My Computer";

            try
            {
                // For drive roots like C:\, show "Local Disk (C:)"
                if (Path.GetPathRoot(path) == path)
                {
                    var driveInfo = new DriveInfo(path);
                    return $"{driveInfo.VolumeLabel} ({path.TrimEnd('\\')}:)".Trim();
                }

                // For regular directories, show just the folder name
                var directoryInfo = new DirectoryInfo(path);
                return directoryInfo.Name;
            }
            catch
            {
                return path;
            }
        }

        /// <summary>
        /// Load bookmarks from registry
        /// </summary>
        private void LoadBookmarks()
        {
            try
            {
                var loadedBookmarks = registryManager.LoadBookmarks();
                bookmarks.Clear();
                bookmarks.AddRange(loadedBookmarks);
                Debug.WriteLine($"Loaded {bookmarks.Count} bookmarks from registry");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error loading bookmarks: {ex.Message}");
            }
        }

        /// <summary>
        /// Save bookmarks to registry
        /// </summary>
        private void SaveBookmarks()
        {
            try
            {
                registryManager.SaveBookmarks(bookmarks);
                Debug.WriteLine($"Saved {bookmarks.Count} bookmarks to registry");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error saving bookmarks: {ex.Message}");
            }
        }

        /// <summary>
        /// Populate a menu with bookmark items
        /// </summary>
        public void PopulateFavoritesMenu(ToolStripMenuItem favoritesMenu, Action<string> navigateAction)
        {
            // Clear existing items except the "Add to Favorites" item
            var itemsToRemove = favoritesMenu.DropDownItems.Cast<ToolStripItem>()
                .Where(item => item.Tag?.ToString() == "bookmark")
                .ToList();
            
            foreach (var item in itemsToRemove)
            {
                favoritesMenu.DropDownItems.Remove(item);
            }

            // Add bookmark items
            var sortedBookmarks = GetBookmarks();
            if (sortedBookmarks.Count > 0)
            {
                // Add separator if there are other items
                if (favoritesMenu.DropDownItems.Count > 0)
                {
                    var separator = new ToolStripSeparator { Tag = "bookmark" };
                    favoritesMenu.DropDownItems.Add(separator);
                }

                // Add bookmark menu items
                foreach (var bookmark in sortedBookmarks)
                {
                    var menuItem = new ToolStripMenuItem(bookmark.Name)
                    {
                        Tag = "bookmark",
                        ToolTipText = bookmark.Path
                    };
                    
                    menuItem.Click += (s, e) => navigateAction(bookmark.Path);
                    favoritesMenu.DropDownItems.Add(menuItem);
                }
            }
        }
    }
}