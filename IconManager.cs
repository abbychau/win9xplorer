using System.Diagnostics;
using System.Drawing.Drawing2D;

namespace win9xplorer
{
    /// <summary>
    /// Manages icon loading and creation for the file explorer
    /// </summary>
    internal class IconManager
    {
        private readonly ImageList imageListSmall;
        private readonly ImageList imageListLarge;

        public IconManager(ImageList smallImageList, ImageList largeImageList)
        {
            imageListSmall = smallImageList;
            imageListLarge = largeImageList;
        }

        public void SetupImageLists()
        {
            // Clear existing images
            imageListSmall.Images.Clear();
            imageListLarge.Images.Clear();
            
            // Add system icons for common items
            var folderIconSmall = GetSystemIcon("folder", true, false);
            var driveIconSmall = GetSystemIcon("C:\\", true, false);
            var computerIconSmall = GetSystemIcon("", true, false, true);
            
            var folderIconLarge = GetSystemIcon("folder", false, false);
            var driveIconLarge = GetSystemIcon("C:\\", false, false);
            var computerIconLarge = GetSystemIcon("", false, false, true);
            
            if (folderIconSmall != null) imageListSmall.Images.Add("folder", folderIconSmall);
            if (driveIconSmall != null) imageListSmall.Images.Add("drive", driveIconSmall);
            if (computerIconSmall != null) imageListSmall.Images.Add("computer", computerIconSmall);
            
            if (folderIconLarge != null) imageListLarge.Images.Add("folder", folderIconLarge);
            if (driveIconLarge != null) imageListLarge.Images.Add("drive", driveIconLarge);
            if (computerIconLarge != null) imageListLarge.Images.Add("computer", computerIconLarge);
        }

        public Icon? GetSystemIcon(string filePath, bool smallIcon, bool isFile, bool isComputer = false)
        {
            WinApi.SHFILEINFO shfi = new WinApi.SHFILEINFO();
            uint flags = WinApi.SHGFI_ICON | (smallIcon ? WinApi.SHGFI_SMALLICON : WinApi.SHGFI_LARGEICON);
            
            if (isComputer)
            {
                // Get "My Computer" icon
                flags |= WinApi.SHGFI_USEFILEATTRIBUTES;
                WinApi.SHGetFileInfo("", 0, ref shfi, (uint)System.Runtime.InteropServices.Marshal.SizeOf(shfi), flags | WinApi.SHGFI_PIDL);
            }
            else if (!isFile)
            {
                // For folders
                flags |= WinApi.SHGFI_USEFILEATTRIBUTES;
                WinApi.SHGetFileInfo(filePath, WinApi.FILE_ATTRIBUTE_DIRECTORY, ref shfi, (uint)System.Runtime.InteropServices.Marshal.SizeOf(shfi), flags);
            }
            else
            {
                // For files
                WinApi.SHGetFileInfo(filePath, 0, ref shfi, (uint)System.Runtime.InteropServices.Marshal.SizeOf(shfi), flags);
            }

            Icon? icon = null;
            if (shfi.hIcon != IntPtr.Zero)
            {
                icon = Icon.FromHandle(shfi.hIcon).Clone() as Icon;
                WinApi.DestroyIcon(shfi.hIcon);
            }

            return icon ?? SystemIcons.WinLogo;
        }

        public string GetIconKey(string filePath, bool isDirectory)
        {
            if (isDirectory)
            {
                return "folder";
            }

            string extension = Path.GetExtension(filePath).ToLower();
            string iconKey = $"file_{extension}";

            // Check if we already have this icon in our image lists
            if (!imageListSmall.Images.ContainsKey(iconKey))
            {
                // Get the system icon for this file type
                Icon? smallIcon = GetSystemIcon(filePath, true, true);
                Icon? largeIcon = GetSystemIcon(filePath, false, true);

                if (smallIcon != null)
                {
                    imageListSmall.Images.Add(iconKey, smallIcon);
                }
                if (largeIcon != null)
                {
                    imageListLarge.Images.Add(iconKey, largeIcon);
                }
            }

            return iconKey;
        }

        public void SetupToolbarIcons(ToolStripButton btnBack, ToolStripButton btnForward, ToolStripButton btnUp, ToolStripButton btnRefresh,
            ToolStripButton btnViewLargeIcons, ToolStripButton btnViewSmallIcons, ToolStripButton btnViewList, ToolStripButton btnViewDetails, int iconSize = 16)
        {
            Debug.WriteLine("=== Starting SetupToolbarIcons ===");
            
            // First try embedded Windows XP icons, then file-system fallback
            try
            {
                string? iconsFolder = FindWindowsXPIconsFolder();
                if (string.IsNullOrEmpty(iconsFolder))
                {
                    Debug.WriteLine("Windows XP Icons folder not found. Embedded resources only.");
                }

                // Navigation icons using Windows XP style icons with multiple filename attempts
                btnBack.Image = LoadWindowsXPIconWithAlternatives(iconsFolder, new[] { "Back.png" }, "Back", iconSize) 
                    ?? GetBestShellIcon(new int[] { 149, 126, 0 }, "Back", true);
                    
                btnForward.Image = LoadWindowsXPIconWithAlternatives(iconsFolder, new[] { "Forward.png" }, "Forward", iconSize) 
                    ?? GetBestShellIcon(new int[] { 150, 127, 1 }, "Forward", true);
                    
                btnUp.Image = LoadWindowsXPIconWithAlternatives(iconsFolder, new[] { "Up.png" }, "Up", iconSize) 
                    ?? GetBestShellIcon(new int[] { 146, 25, 5 }, "Up", true);
                    
                btnRefresh.Image = LoadWindowsXPIconWithAlternatives(iconsFolder, new[] { "Refresh.png", "IE Refresh.png" }, "Refresh", iconSize) 
                    ?? GetBestShellIcon(new int[] { 238, 239, 147 }, "Refresh", true);
                
                // View As icons using Windows XP style icons with better mappings and alternatives
                btnViewLargeIcons.Image = LoadWindowsXPIconWithAlternatives(iconsFolder, new[] { "Icon View.png", "Tile View.png" }, "LargeIcons", iconSize) 
                    ?? GetBestShellIcon(new int[] { 269, 3, 2 }, "LargeIcons", true);
                    
                btnViewSmallIcons.Image = LoadWindowsXPIconWithAlternatives(iconsFolder, new[] { "Folder View - Classic.png", "Thumbnail View.png" }, "SmallIcons", iconSize) 
                    ?? GetBestShellIcon(new int[] { 268, 152, 2 }, "SmallIcons", true);
                    
                btnViewList.Image = LoadWindowsXPIconWithAlternatives(iconsFolder, new[] { "Checklist.png", "List View.png" }, "List", iconSize) 
                    ?? GetBestShellIcon(new int[] { 267, 151, 1 }, "List", true);
                    
                btnViewDetails.Image = LoadWindowsXPIconWithAlternatives(iconsFolder, new[] { "Detail View.png", "Details.png" }, "Details", iconSize) 
                    ?? GetBestShellIcon(new int[] { 266, 153, 4 }, "Details", true);
                
                Debug.WriteLine("Windows XP Icons setup completed");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to load Windows XP icons: {ex.Message}");
                // Fallback to shell icons
                SetupShellIconsToolbar(btnBack, btnForward, btnUp, btnRefresh, btnViewLargeIcons, btnViewSmallIcons, btnViewList, btnViewDetails);
            }
            
            Debug.WriteLine("=== SetupToolbarIcons Completed ===");
        }

        private Bitmap? LoadWindowsXPIconWithAlternatives(string? iconsFolder, string[] iconFileNames, string buttonName, int size = 16)
        {
            foreach (string iconFileName in iconFileNames)
            {
                var bitmap = LoadWindowsXPIcon(iconsFolder, iconFileName, size);
                if (bitmap != null)
                {
                    Debug.WriteLine($"Successfully loaded {buttonName} icon from: {iconFileName}");
                    return bitmap;
                }
            }
            
            Debug.WriteLine($"Failed to load {buttonName} icon from any of the attempted filenames");
            return null;
        }

        private Bitmap? LoadWindowsXPIcon(string? iconsFolder, string iconFileName, int size = 16)
        {
            try
            {
                var embedded = LoadWindowsXPIconFromEmbedded(iconFileName, size);
                if (embedded != null)
                {
                    return embedded;
                }

                if (string.IsNullOrWhiteSpace(iconsFolder))
                {
                    return null;
                }

                string iconsPath = Path.Combine(iconsFolder, iconFileName);
                
                // Check if the icon file exists
                if (!File.Exists(iconsPath))
                {
                    Debug.WriteLine($"Windows XP icon not found: {iconsPath}");
                    return null;
                }

                // Check file size to make sure it's not corrupt
                var fileInfo = new FileInfo(iconsPath);
                if (fileInfo.Length == 0)
                {
                    Debug.WriteLine($"Windows XP icon file is empty: {iconsPath}");
                    return null;
                }

                Debug.WriteLine($"Loading icon: {iconsPath} (Size: {fileInfo.Length} bytes)");

                // Load the image and resize for toolbar usage
                using (var originalImage = Image.FromFile(iconsPath))
                {
                    Debug.WriteLine($"Original image size: {originalImage.Width}x{originalImage.Height}, Format: {originalImage.PixelFormat}");
                    
                    var bitmap = new Bitmap(size, size, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                    using (var g = Graphics.FromImage(bitmap))
                    {
                        g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                        g.SmoothingMode = SmoothingMode.HighQuality;
                        g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                        g.CompositingQuality = CompositingQuality.HighQuality;
                        g.CompositingMode = CompositingMode.SourceOver;
                        
                        // Clear background to transparent
                        g.Clear(Color.Transparent);
                        
                        // Draw the resized icon with proper transparency
                        g.DrawImage(originalImage, new Rectangle(0, 0, size, size), 
                            new Rectangle(0, 0, originalImage.Width, originalImage.Height), GraphicsUnit.Pixel);
                    }
                    
                    Debug.WriteLine($"Successfully created {size}x{size} bitmap from: {iconFileName}");
                    return bitmap;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error loading Windows XP icon {iconFileName}: {ex.Message}");
                Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                return null;
            }
        }

        private static Bitmap? LoadWindowsXPIconFromEmbedded(string iconFileName, int size)
        {
            try
            {
                var assembly = typeof(IconManager).Assembly;
                var resourceName = assembly
                    .GetManifestResourceNames()
                    .FirstOrDefault(name =>
                        name.EndsWith(iconFileName, StringComparison.OrdinalIgnoreCase) &&
                        (name.Contains("Windows XP Icons", StringComparison.OrdinalIgnoreCase) ||
                         name.Contains("Windows_XP_Icons", StringComparison.OrdinalIgnoreCase)));

                if (string.IsNullOrWhiteSpace(resourceName))
                {
                    return null;
                }

                using var stream = assembly.GetManifestResourceStream(resourceName);
                if (stream == null)
                {
                    return null;
                }

                using var originalImage = Image.FromStream(stream);
                var bitmap = new Bitmap(size, size, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                using var g = Graphics.FromImage(bitmap);
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.SmoothingMode = SmoothingMode.HighQuality;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.CompositingQuality = CompositingQuality.HighQuality;
                g.CompositingMode = CompositingMode.SourceOver;
                g.Clear(Color.Transparent);
                g.DrawImage(originalImage, new Rectangle(0, 0, size, size), new Rectangle(0, 0, originalImage.Width, originalImage.Height), GraphicsUnit.Pixel);
                return bitmap;
            }
            catch
            {
                return null;
            }
        }

        private void SetupShellIconsToolbar(ToolStripButton btnBack, ToolStripButton btnForward, ToolStripButton btnUp, ToolStripButton btnRefresh,
            ToolStripButton btnViewLargeIcons, ToolStripButton btnViewSmallIcons, ToolStripButton btnViewList, ToolStripButton btnViewDetails)
        {
            // Fallback to original shell32 icons if Windows XP icons are not available
            try
            {
                // Navigation icons
                btnBack.Image = GetBestShellIcon(new int[] { 149, 126, 0 }, "Back", true);        // Better back arrow
                btnForward.Image = GetBestShellIcon(new int[] { 150, 127, 1 }, "Forward", true);  // Better forward arrow
                btnUp.Image = GetBestShellIcon(new int[] { 146, 25, 5 }, "Up", true);             // Better up arrow
                btnRefresh.Image = GetBestShellIcon(new int[] { 238, 239, 147 }, "Refresh", true); // Better refresh icon
                
                // View As icons - using authentic Windows 95/98 view mode icons
                btnViewLargeIcons.Image = GetBestShellIcon(new int[] { 269, 3, 2 }, "LargeIcons", true);   // Large icons view
                btnViewSmallIcons.Image = GetBestShellIcon(new int[] { 268, 152, 2 }, "SmallIcons", true); // Small icons view
                btnViewList.Image = GetBestShellIcon(new int[] { 267, 151, 1 }, "List", true);             // List view
                btnViewDetails.Image = GetBestShellIcon(new int[] { 266, 153, 4 }, "Details", true);       // Details view
            }
            catch (Exception ex)
            {
                // If shell icon extraction fails, use enhanced fallback icons
                Debug.WriteLine($"Failed to load shell icons: {ex.Message}");
                SetupFallbackToolbarIcons(btnBack, btnForward, btnUp, btnRefresh, btnViewLargeIcons, btnViewSmallIcons, btnViewList, btnViewDetails);
            }
        }

        private void SetupFallbackToolbarIcons(ToolStripButton btnBack, ToolStripButton btnForward, ToolStripButton btnUp, ToolStripButton btnRefresh,
            ToolStripButton btnViewLargeIcons, ToolStripButton btnViewSmallIcons, ToolStripButton btnViewList, ToolStripButton btnViewDetails)
        {
            // Fallback to real PNG assets
            btnBack.Image = LoadFallbackToolbarAsset("Back", 16) ?? SystemIcons.Application.ToBitmap();
            btnForward.Image = LoadFallbackToolbarAsset("Forward", 16) ?? SystemIcons.Application.ToBitmap();
            btnUp.Image = LoadFallbackToolbarAsset("Up", 16) ?? SystemIcons.Application.ToBitmap();
            btnRefresh.Image = LoadFallbackToolbarAsset("Refresh", 16) ?? SystemIcons.Application.ToBitmap();
            
            // Fallback icons for view modes
            btnViewLargeIcons.Image = CreateViewIcon("LargeIcons", 16);
            btnViewSmallIcons.Image = CreateViewIcon("SmallIcons", 16);
            btnViewList.Image = CreateViewIcon("List", 16);
            btnViewDetails.Image = CreateViewIcon("Details", 16);
        }

        public Bitmap CreateViewIcon(string viewType, int size)
        {
            string[] candidates = viewType switch
            {
                "LargeIcons" => new[] { "Icon View.png", "Tile View.png" },
                "SmallIcons" => new[] { "Folder View - Classic.png", "Thumbnail View.png" },
                "List" => new[] { "Checklist.png", "List View.png" },
                "Details" => new[] { "Detail View.png", "Details.png" },
                "folders" => new[] { "Folder View - Common Tasks.png", "Folder View - Classic.png", "Folder Closed.png" },
                "star" => new[] { "Favorites.png", "Add Network Place.png", "Address Book.png" },
                "navigate" => new[] { "Go.png", "Open.png", "Double Click.png" },
                _ => Array.Empty<string>()
            };

            string? iconsFolder = FindWindowsXPIconsFolder();
            if (candidates.Length > 0)
            {
                Bitmap? loaded = LoadWindowsXPIconWithAlternatives(iconsFolder, candidates, viewType, size);
                if (loaded != null)
                {
                    return loaded;
                }
            }

            return ResizeImage(SystemIcons.Application.ToBitmap(), size);
        }

        private static Bitmap ResizeImage(Image source, int size)
        {
            var bitmap = new Bitmap(size, size, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using var graphics = Graphics.FromImage(bitmap);
            graphics.Clear(Color.Transparent);
            graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
            graphics.SmoothingMode = SmoothingMode.HighQuality;
            graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
            graphics.DrawImage(source, new Rectangle(0, 0, size, size));
            return bitmap;
        }

        private Bitmap GetBestShellIcon(int[] iconIndices, string buttonName, bool smallIcon = true)
        {
            // Try each icon index until we find one that works
            foreach (int iconIndex in iconIndices)
            {
                try
                {
                    IntPtr hIcon = WinApi.ExtractIcon(Process.GetCurrentProcess().Handle, "shell32.dll", iconIndex);
                    
                    if (hIcon != IntPtr.Zero && hIcon != (IntPtr)1)
                    {
                        try
                        {
                            using (Icon icon = Icon.FromHandle(hIcon))
                            {
                                int size = smallIcon ? 16 : 32;
                                Bitmap bitmap = new Bitmap(size, size);
                                using (Graphics g = Graphics.FromImage(bitmap))
                                {
                                    g.Clear(Color.Transparent);
                                    // Use high-quality rendering for better icon appearance
                                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                                    g.SmoothingMode = SmoothingMode.HighQuality;
                                    g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                                    g.DrawIcon(icon, new Rectangle(0, 0, size, size));
                                }
                                Debug.WriteLine($"{buttonName} button using shell32.dll icon index {iconIndex}");
                                return bitmap;
                            }
                        }
                        finally
                        {
                            WinApi.DestroyIcon(hIcon);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Failed to extract icon {iconIndex} for {buttonName}: {ex.Message}");
                    continue;
                }
            }
            
            // If all shell32 icons fail, try comctl32.dll (IE4 style icons)
            var comctl32Icon = GetComctl32Icon(buttonName, smallIcon);
            if (comctl32Icon != null)
                return comctl32Icon;
            
            // Final fallback to custom icons
            return buttonName switch
            {
                "LargeIcons" or "SmallIcons" or "List" or "Details" => CreateViewIcon(buttonName, smallIcon ? 16 : 32),
                _ => LoadFallbackToolbarAsset(buttonName, smallIcon ? 16 : 32) ?? SystemIcons.Application.ToBitmap()
            };
        }

        private Bitmap? GetComctl32Icon(string buttonName, bool smallIcon = true)
        {
            try
            {
                // IE4 introduced new toolbar icons in comctl32.dll
                string dllPath = Path.Combine(Environment.SystemDirectory, "comctl32.dll");
                if (!File.Exists(dllPath)) return null;

                int iconIndex = buttonName switch
                {
                    "Back" => 0,    // Standard back arrow
                    "Forward" => 1, // Standard forward arrow
                    "Up" => 2,      // Standard up arrow
                    "Refresh" => 5, // Standard refresh
                    _ => -1
                };

                if (iconIndex >= 0)
                {
                    IntPtr hIcon = WinApi.ExtractIcon(Process.GetCurrentProcess().Handle, dllPath, iconIndex);
                    
                    if (hIcon != IntPtr.Zero && hIcon != (IntPtr)1)
                    {
                        try
                        {
                            using (Icon icon = Icon.FromHandle(hIcon))
                            {
                                int size = smallIcon ? 16 : 32;
                                Bitmap bitmap = new Bitmap(size, size);
                                using (Graphics g = Graphics.FromImage(bitmap))
                                {
                                    g.Clear(Color.Transparent);
                                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                                    g.SmoothingMode = SmoothingMode.HighQuality;
                                    g.DrawIcon(icon, new Rectangle(0, 0, size, size));
                                }
                                Debug.WriteLine($"{buttonName} button using comctl32.dll icon index {iconIndex}");
                                return bitmap;
                            }
                        }
                        finally
                        {
                            WinApi.DestroyIcon(hIcon);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to extract comctl32 icon for {buttonName}: {ex.Message}");
            }
            
            return null;
        }

        private Bitmap? LoadFallbackToolbarAsset(string iconType, int size)
        {
            string[] candidates = iconType switch
            {
                "Back" => new[] { "Back.png" },
                "Forward" => new[] { "Forward.png" },
                "Up" => new[] { "Up.png" },
                "Refresh" => new[] { "Refresh.png", "IE Refresh.png" },
                _ => Array.Empty<string>()
            };

            if (candidates.Length == 0)
            {
                return null;
            }

            string? iconsFolder = FindWindowsXPIconsFolder();
            return LoadWindowsXPIconWithAlternatives(iconsFolder, candidates, iconType, size);
        }

        private string? FindWindowsXPIconsFolder()
        {
            string[] possiblePaths =
            {
                Path.Combine(Application.StartupPath, "Windows XP Icons"),
                Path.Combine(Directory.GetCurrentDirectory(), "Windows XP Icons"),
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Windows XP Icons")
            };

            foreach (string path in possiblePaths)
            {
                if (Directory.Exists(path))
                {
                    Debug.WriteLine($"Found Windows XP Icons folder at: {path}");
                    return path;
                }
            }
            return null;
        }

        public Icon LoadApplicationIcon()
        {
            string[] appIconCandidates =
            {
                "Explorer.png",
                "File Explorer.png",
                "Application Window.png",
                "Explorer Delete.png",
                "Desktop.png"
            };

            string? iconsFolder = FindWindowsXPIconsFolder();
            foreach (string iconFileName in appIconCandidates)
            {
                using var sourceImage = LoadWindowsXPIcon(iconsFolder, iconFileName, 32);
                if (sourceImage == null)
                {
                    continue;
                }

                using var bitmap = new Bitmap(32, 32, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
                using (var graphics = Graphics.FromImage(bitmap))
                {
                    graphics.Clear(Color.Transparent);
                    graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    graphics.SmoothingMode = SmoothingMode.HighQuality;
                    graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
                    graphics.DrawImage(sourceImage, new Rectangle(0, 0, 32, 32));
                }

                IntPtr hIcon = bitmap.GetHicon();
                try
                {
                    using var temporaryIcon = Icon.FromHandle(hIcon);
                    return (Icon)temporaryIcon.Clone();
                }
                finally
                {
                    WinApi.DestroyIcon(hIcon);
                }
            }

            return SystemIcons.Application;
        }

        /// <summary>
        /// Test method to verify Windows XP icons are loading - can be called for debugging
        /// </summary>
        public void TestWindowsXPIconLoading()
        {
            Debug.WriteLine("=== Testing Windows XP Icon Loading ===");
            
            string iconsFolder = "";
            string[] possiblePaths = {
                Path.Combine(Application.StartupPath, "Windows XP Icons"),
                Path.Combine(Directory.GetCurrentDirectory(), "Windows XP Icons"),
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Windows XP Icons")
            };

            foreach (var path in possiblePaths)
            {
                if (Directory.Exists(path))
                {
                    iconsFolder = path;
                    Debug.WriteLine($"Using icons folder: {path}");
                    break;
                }
            }

            if (string.IsNullOrEmpty(iconsFolder))
            {
                Debug.WriteLine("ERROR: No Windows XP Icons folder found!");
                MessageBox.Show("Windows XP Icons folder not found!", "Icon Test", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Test loading each icon
            string[] testIcons = { 
                "Back.png", "Forward.png", "Up.png", "IE Refresh.png", 
                "Detail View.png", "Icon View.png", "Checklist.png", "Tile View.png" 
            };

            int successCount = 0;
            foreach (var iconName in testIcons)
            {
                var bitmap = LoadWindowsXPIcon(iconsFolder, iconName);
                if (bitmap != null)
                {
                    Debug.WriteLine($"? Successfully loaded: {iconName}");
                    successCount++;
                    bitmap.Dispose(); // Clean up test bitmap
                }
                else
                {
                    Debug.WriteLine($"? Failed to load: {iconName}");
                }
            }

            string message = $"Icon Loading Test Results:\n{successCount}/{testIcons.Length} icons loaded successfully\n\nCheck Debug Output for details.";
            MessageBox.Show(message, "Windows XP Icon Test", MessageBoxButtons.OK, 
                successCount == testIcons.Length ? MessageBoxIcon.Information : MessageBoxIcon.Warning);
            
            Debug.WriteLine($"=== Test Complete: {successCount}/{testIcons.Length} icons loaded ===");
        }
    }
}
