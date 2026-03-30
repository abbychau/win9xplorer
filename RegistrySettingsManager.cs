using Microsoft.Win32;
using System.Drawing;
using System.Windows.Forms;
using System.Diagnostics;

namespace win9xplorer
{
    /// <summary>
    /// Manages application settings persistence using Windows Registry
    /// </summary>
    internal class RegistrySettingsManager
    {
        private const string REGISTRY_PATH = @"SOFTWARE\win9xplorer";
        private const string BOOKMARKS_KEY = "Bookmarks";
        private const string WINDOW_KEY = "Window";
        private const string TOOLSTRIP_KEY = "ToolStrip";
        private const string FONTS_KEY = "Fonts";
        private const string FILE_OPERATIONS_KEY = "FileOperations";

        public class WindowSettings
        {
            public Size Size { get; set; } = new Size(800, 600);
            public Point Location { get; set; } = new Point(100, 100);
            public FormWindowState WindowState { get; set; } = FormWindowState.Normal;
            public int SplitterDistance { get; set; } = 200;
            public bool TreeViewVisible { get; set; } = true;
            public bool ToolbarVisible { get; set; } = true;
            public bool StatusBarVisible { get; set; } = true;
            public string LastPath { get; set; } = "My Computer";
            public View ListViewMode { get; set; } = View.Details;
        }

        public class ToolStripSettings
        {
            public Point ToolStripLocation { get; set; } = new Point(0, 0);
            public Point AddressStripLocation { get; set; } = new Point(0, 25);
            public Size AddressTextBoxSize { get; set; } = new Size(300, 27);
        }

        public class FontSettings
        {
            public Font TreeViewFont { get; set; } = SystemFonts.DefaultFont;
            public Font ListViewFont { get; set; } = SystemFonts.DefaultFont;
        }

        public class FileOperationSettings
        {
            public FileConflictStrategy ConflictStrategy { get; set; } = FileConflictStrategy.AskUser;
        }

        /// <summary>
        /// Saves window settings to registry
        /// </summary>
        public void SaveWindowSettings(Form form, SplitContainer splitContainer, ListView listView)
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey($@"{REGISTRY_PATH}\{WINDOW_KEY}"))
                {
                    // Window position and size (only if not minimized or maximized)
                    if (form.WindowState == FormWindowState.Normal)
                    {
                        key.SetValue("Width", form.Width);
                        key.SetValue("Height", form.Height);
                        key.SetValue("Left", form.Left);
                        key.SetValue("Top", form.Top);
                    }
                    
                    key.SetValue("WindowState", (int)form.WindowState);
                    key.SetValue("SplitterDistance", splitContainer.SplitterDistance);
                    key.SetValue("TreeViewVisible", !splitContainer.Panel1Collapsed);
                    key.SetValue("LastPath", GetCurrentPath(form));
                    key.SetValue("ListViewMode", (int)listView.View);
                    
                    Debug.WriteLine("Window settings saved to registry");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error saving window settings: {ex.Message}");
            }
        }

        /// <summary>
        /// Loads window settings from registry
        /// </summary>
        public WindowSettings LoadWindowSettings()
        {
            var settings = new WindowSettings();
            
            try
            {
                using (RegistryKey? key = Registry.CurrentUser.OpenSubKey($@"{REGISTRY_PATH}\{WINDOW_KEY}"))
                {
                    if (key != null)
                    {
                        settings.Size = new Size(
                            (int)(key.GetValue("Width") ?? 871),
                            (int)(key.GetValue("Height") ?? 384)
                        );
                        
                        settings.Location = new Point(
                            (int)(key.GetValue("Left") ?? 100),
                            (int)(key.GetValue("Top") ?? 100)
                        );
                        
                        settings.WindowState = (FormWindowState)(key.GetValue("WindowState") ?? 0);
                        settings.SplitterDistance = (int)(key.GetValue("SplitterDistance") ?? 217);
                        settings.TreeViewVisible = (bool)(key.GetValue("TreeViewVisible") ?? true);
                        settings.LastPath = key.GetValue("LastPath")?.ToString() ?? "My Computer";
                        settings.ListViewMode = (View)((int)(key.GetValue("ListViewMode") ?? 3)); // Default to Details
                        
                        Debug.WriteLine("Window settings loaded from registry");
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error loading window settings: {ex.Message}");
            }
            
            return settings;
        }

        /// <summary>
        /// Saves toolbar positions to registry
        /// </summary>
        public void SaveToolStripSettings(ToolStrip toolStrip, ToolStrip addressStrip, ToolStripTextBox addressTextBox)
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey($@"{REGISTRY_PATH}\{TOOLSTRIP_KEY}"))
                {
                    key.SetValue("ToolStripX", toolStrip.Location.X);
                    key.SetValue("ToolStripY", toolStrip.Location.Y);
                    key.SetValue("AddressStripX", addressStrip.Location.X);
                    key.SetValue("AddressStripY", addressStrip.Location.Y);
                    key.SetValue("AddressTextBoxWidth", addressTextBox.Size.Width);
                    key.SetValue("AddressTextBoxHeight", addressTextBox.Size.Height);
                    
                    Debug.WriteLine("ToolStrip settings saved to registry");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error saving toolbar settings: {ex.Message}");
            }
        }

        /// <summary>
        /// Loads toolbar settings from registry
        /// </summary>
        public ToolStripSettings LoadToolStripSettings()
        {
            var settings = new ToolStripSettings();
            
            try
            {
                using (RegistryKey? key = Registry.CurrentUser.OpenSubKey($@"{REGISTRY_PATH}\{TOOLSTRIP_KEY}"))
                {
                    if (key != null)
                    {
                        settings.ToolStripLocation = new Point(
                            (int)(key.GetValue("ToolStripX") ?? 459),
                            (int)(key.GetValue("ToolStripY") ?? 0)
                        );
                        
                        settings.AddressStripLocation = new Point(
                            (int)(key.GetValue("AddressStripX") ?? 6),
                            (int)(key.GetValue("AddressStripY") ?? 0)
                        );
                        
                        settings.AddressTextBoxSize = new Size(
                            (int)(key.GetValue("AddressTextBoxWidth") ?? 300),
                            (int)(key.GetValue("AddressTextBoxHeight") ?? 27)
                        );
                        
                        Debug.WriteLine("ToolStrip settings loaded from registry");
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error loading toolbar settings: {ex.Message}");
            }
            
            return settings;
        }

        /// <summary>
        /// Saves font settings to registry
        /// </summary>
        public void SaveFontSettings(Font treeViewFont, Font listViewFont)
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey($@"{REGISTRY_PATH}\{FONTS_KEY}"))
                {
                    // TreeView font
                    key.SetValue("TreeViewFontName", treeViewFont.Name);
                    key.SetValue("TreeViewFontSize", treeViewFont.Size);
                    key.SetValue("TreeViewFontStyle", (int)treeViewFont.Style);
                    
                    // ListView font
                    key.SetValue("ListViewFontName", listViewFont.Name);
                    key.SetValue("ListViewFontSize", listViewFont.Size);
                    key.SetValue("ListViewFontStyle", (int)listViewFont.Style);
                    
                    Debug.WriteLine("Font settings saved to registry");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error saving font settings: {ex.Message}");
            }
        }

        /// <summary>
        /// Loads font settings from registry
        /// </summary>
        public FontSettings LoadFontSettings()
        {
            var settings = new FontSettings();
            
            try
            {
                using (RegistryKey? key = Registry.CurrentUser.OpenSubKey($@"{REGISTRY_PATH}\{FONTS_KEY}"))
                {
                    if (key != null)
                    {
                        // TreeView font
                        string treeViewFontName = key.GetValue("TreeViewFontName")?.ToString() ?? SystemFonts.DefaultFont.Name;
                        float treeViewFontSize = Convert.ToSingle(key.GetValue("TreeViewFontSize") ?? SystemFonts.DefaultFont.Size);
                        FontStyle treeViewFontStyle = (FontStyle)((int)(key.GetValue("TreeViewFontStyle") ?? 0));
                        
                        // ListView font
                        string listViewFontName = key.GetValue("ListViewFontName")?.ToString() ?? SystemFonts.DefaultFont.Name;
                        float listViewFontSize = Convert.ToSingle(key.GetValue("ListViewFontSize") ?? SystemFonts.DefaultFont.Size);
                        FontStyle listViewFontStyle = (FontStyle)((int)(key.GetValue("ListViewFontStyle") ?? 0));
                        
                        try
                        {
                            settings.TreeViewFont = new Font(treeViewFontName, treeViewFontSize, treeViewFontStyle);
                            settings.ListViewFont = new Font(listViewFontName, listViewFontSize, listViewFontStyle);
                            Debug.WriteLine("Font settings loaded from registry");
                        }
                        catch (Exception fontEx)
                        {
                            Debug.WriteLine($"Error creating fonts from registry: {fontEx.Message}");
                            // Fall back to system fonts if font creation fails
                            settings.TreeViewFont = new Font(SystemFonts.DefaultFont, SystemFonts.DefaultFont.Style);
                            settings.ListViewFont = new Font(SystemFonts.DefaultFont, SystemFonts.DefaultFont.Style);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error loading font settings: {ex.Message}");
            }
            
            return settings;
        }

        public void SaveFileOperationSettings(FileConflictStrategy conflictStrategy)
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey($@"{REGISTRY_PATH}\{FILE_OPERATIONS_KEY}"))
                {
                    key.SetValue("ConflictStrategy", (int)conflictStrategy);
                    Debug.WriteLine($"File operation settings saved: {conflictStrategy}");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error saving file operation settings: {ex.Message}");
            }
        }

        public FileOperationSettings LoadFileOperationSettings()
        {
            var settings = new FileOperationSettings();

            try
            {
                using (RegistryKey? key = Registry.CurrentUser.OpenSubKey($@"{REGISTRY_PATH}\{FILE_OPERATIONS_KEY}"))
                {
                    if (key != null)
                    {
                        int strategyValue = (int)(key.GetValue("ConflictStrategy") ?? (int)FileConflictStrategy.AskUser);
                        if (Enum.IsDefined(typeof(FileConflictStrategy), strategyValue))
                        {
                            settings.ConflictStrategy = (FileConflictStrategy)strategyValue;
                        }

                        Debug.WriteLine($"File operation settings loaded: {settings.ConflictStrategy}");
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error loading file operation settings: {ex.Message}");
            }

            return settings;
        }

        /// <summary>
        /// Saves bookmarks/favorites to registry
        /// </summary>
        public void SaveBookmarks(List<BookmarkManager.Bookmark> bookmarks)
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey($@"{REGISTRY_PATH}\{BOOKMARKS_KEY}"))
                {
                    // Clear existing bookmarks
                    string[] existingKeys = key.GetSubKeyNames();
                    foreach (string keyName in existingKeys)
                    {
                        key.DeleteSubKey(keyName);
                    }
                    
                    // Save new bookmarks
                    for (int i = 0; i < bookmarks.Count; i++)
                    {
                        using (RegistryKey bookmarkKey = key.CreateSubKey($"Bookmark_{i}"))
                        {
                            bookmarkKey.SetValue("Name", bookmarks[i].Name);
                            bookmarkKey.SetValue("Path", bookmarks[i].Path);
                            bookmarkKey.SetValue("DateAdded", bookmarks[i].DateAdded.ToBinary());
                        }
                    }
                    
                    Debug.WriteLine($"Saved {bookmarks.Count} bookmarks to registry");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error saving bookmarks: {ex.Message}");
            }
        }

        /// <summary>
        /// Loads bookmarks/favorites from registry
        /// </summary>
        public List<BookmarkManager.Bookmark> LoadBookmarks()
        {
            var bookmarks = new List<BookmarkManager.Bookmark>();
            
            try
            {
                using (RegistryKey? key = Registry.CurrentUser.OpenSubKey($@"{REGISTRY_PATH}\{BOOKMARKS_KEY}"))
                {
                    if (key != null)
                    {
                        string[] bookmarkKeys = key.GetSubKeyNames();
                        
                        foreach (string keyName in bookmarkKeys.OrderBy(k => k))
                        {
                            using (RegistryKey? bookmarkKey = key.OpenSubKey(keyName))
                            {
                                if (bookmarkKey != null)
                                {
                                    string name = bookmarkKey.GetValue("Name")?.ToString() ?? "";
                                    string path = bookmarkKey.GetValue("Path")?.ToString() ?? "";
                                    long dateAddedBinary = Convert.ToInt64(bookmarkKey.GetValue("DateAdded") ?? DateTime.Now.ToBinary());
                                    DateTime dateAdded = DateTime.FromBinary(dateAddedBinary);
                                    
                                    if (!string.IsNullOrEmpty(name) && !string.IsNullOrEmpty(path))
                                    {
                                        bookmarks.Add(new BookmarkManager.Bookmark(name, path, dateAdded));
                                    }
                                }
                            }
                        }
                        
                        Debug.WriteLine($"Loaded {bookmarks.Count} bookmarks from registry");
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error loading bookmarks: {ex.Message}");
            }
            
            return bookmarks;
        }

        /// <summary>
        /// Applies window settings to the form
        /// </summary>
        public void ApplyWindowSettings(Form form, SplitContainer splitContainer, ListView listView, WindowSettings settings)
        {
            try
            {
                // Validate screen bounds before applying position
                Rectangle screenBounds = Screen.PrimaryScreen?.WorkingArea ?? Screen.FromControl(form).WorkingArea;
                
                // Ensure the window is at least partially visible
                Point location = settings.Location;
                if (location.X < screenBounds.Left - 50 || location.X > screenBounds.Right - 50)
                    location.X = 100;
                if (location.Y < screenBounds.Top - 50 || location.Y > screenBounds.Bottom - 50)
                    location.Y = 100;
                
                // Ensure minimum size
                Size size = settings.Size;
                if (size.Width < 400) size.Width = 400;
                if (size.Height < 300) size.Height = 300;
                
                form.StartPosition = FormStartPosition.Manual;
                form.Location = location;
                form.Size = size;
                
                // Set window state last
                form.WindowState = settings.WindowState;
                
                // Apply splitter distance with bounds checking
                int splitterDistance = Math.Max(100, Math.Min(settings.SplitterDistance, splitContainer.Width - 200));
                splitContainer.SplitterDistance = splitterDistance;
                
                // Apply tree view visibility
                splitContainer.Panel1Collapsed = !settings.TreeViewVisible;
                
                // Apply list view mode
                listView.View = settings.ListViewMode;
                
                Debug.WriteLine($"Applied window settings - Size: {size}, Location: {location}, State: {settings.WindowState}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error applying window settings: {ex.Message}");
            }
        }

        /// <summary>
        /// Applies toolbar settings to toolstrips
        /// </summary>
        public void ApplyToolStripSettings(ToolStrip toolStrip, ToolStrip addressStrip, ToolStripTextBox addressTextBox, ToolStripSettings settings)
        {
            try
            {
                // Apply toolbar positions
                toolStrip.Location = settings.ToolStripLocation;
                addressStrip.Location = settings.AddressStripLocation;
                
                // Apply address textbox size with bounds checking
                Size textBoxSize = settings.AddressTextBoxSize;
                if (textBoxSize.Width < 150) textBoxSize.Width = 150;
                if (textBoxSize.Width > 500) textBoxSize.Width = 500;
                
                addressTextBox.Size = textBoxSize;
                
                Debug.WriteLine($"Applied toolbar settings - ToolStrip: {settings.ToolStripLocation}, AddressStrip: {settings.AddressStripLocation}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error applying toolbar settings: {ex.Message}");
            }
        }

        /// <summary>
        /// Applies font settings to controls
        /// </summary>
        public void ApplyFontSettings(TreeView treeView, ListView listView, FontSettings settings)
        {
            try
            {
                treeView.Font = new Font(settings.TreeViewFont, settings.TreeViewFont.Style);
                listView.Font = new Font(settings.ListViewFont, settings.ListViewFont.Style);
                
                Debug.WriteLine($"Applied font settings - TreeView: {settings.TreeViewFont.Name}, ListView: {settings.ListViewFont.Name}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error applying font settings: {ex.Message}");
            }
        }

        /// <summary>
        /// Gets the current path from the form for saving
        /// </summary>
        private string GetCurrentPath(Form form)
        {
            // Try to get the address textbox value through reflection
            try
            {
                var addressTextBox = FindControl<ToolStripTextBox>(form, "txtAddress");
                return addressTextBox?.Text ?? "My Computer";
            }
            catch
            {
                return "My Computer";
            }
        }

        /// <summary>
        /// Helper method to find controls by name recursively
        /// </summary>
        private T? FindControl<T>(Control parent, string name) where T : class
        {
            foreach (Control control in parent.Controls)
            {
                if (control.Name == name && control is T)
                    return control as T;
                
                // Check if it's a container control
                if (control.HasChildren)
                {
                    var found = FindControl<T>(control, name);
                    if (found != null)
                        return found;
                }
                
                // Special handling for ToolStripContainer and ToolStrip
                if (control is ToolStripContainer container)
                {
                    foreach (ToolStrip strip in container.TopToolStripPanel.Controls.OfType<ToolStrip>())
                    {
                        foreach (ToolStripItem item in strip.Items)
                        {
                            if (item.Name == name && item is T)
                                return item as T;
                        }
                    }
                }
            }
            
            return null;
        }

        /// <summary>
        /// Clears all registry settings (for testing or reset purposes)
        /// </summary>
        public void ClearAllSettings()
        {
            try
            {
                Registry.CurrentUser.DeleteSubKeyTree(REGISTRY_PATH, false);
                Debug.WriteLine("All registry settings cleared");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error clearing registry settings: {ex.Message}");
            }
        }
    }
}
