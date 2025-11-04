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
            ToolStripButton btnViewLargeIcons, ToolStripButton btnViewSmallIcons, ToolStripButton btnViewList, ToolStripButton btnViewDetails)
        {
            Debug.WriteLine("=== Starting SetupToolbarIcons ===");
            
            // First try to load Windows XP icons from the "Windows XP Icons" folder
            try
            {
                // Try multiple possible icon folder locations
                string[] possiblePaths = {
                    Path.Combine(Application.StartupPath, "Windows XP Icons"),
                    Path.Combine(Directory.GetCurrentDirectory(), "Windows XP Icons"),
                    Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Windows XP Icons")
                };

                string iconsFolder = "";
                foreach (var path in possiblePaths)
                {
                    if (Directory.Exists(path))
                    {
                        iconsFolder = path;
                        Debug.WriteLine($"Found Windows XP Icons folder at: {path}");
                        break;
                    }
                }

                if (string.IsNullOrEmpty(iconsFolder))
                {
                    Debug.WriteLine("Windows XP Icons folder not found in any expected location");
                    SetupShellIconsToolbar(btnBack, btnForward, btnUp, btnRefresh, btnViewLargeIcons, btnViewSmallIcons, btnViewList, btnViewDetails);
                    return;
                }

                // Navigation icons using Windows XP style icons with multiple filename attempts
                btnBack.Image = LoadWindowsXPIconWithAlternatives(iconsFolder, new[] { "Back.png", "back.png", "Back.ico" }, "Back") 
                    ?? GetBestShellIcon(new int[] { 149, 126, 0 }, "Back", true);
                    
                btnForward.Image = LoadWindowsXPIconWithAlternatives(iconsFolder, new[] { "Forward.png", "forward.png", "Forward.ico" }, "Forward") 
                    ?? GetBestShellIcon(new int[] { 150, 127, 1 }, "Forward", true);
                    
                btnUp.Image = LoadWindowsXPIconWithAlternatives(iconsFolder, new[] { "Up.png", "up.png", "Up.ico" }, "Up") 
                    ?? GetBestShellIcon(new int[] { 146, 25, 5 }, "Up", true);
                    
                btnRefresh.Image = LoadWindowsXPIconWithAlternatives(iconsFolder, new[] { "IE Refresh.png", "Refresh.png", "refresh.png", "Refresh.ico" }, "Refresh") 
                    ?? GetBestShellIcon(new int[] { 238, 239, 147 }, "Refresh", true);
                
                // View As icons using Windows XP style icons with better mappings and alternatives
                btnViewLargeIcons.Image = LoadWindowsXPIconWithAlternatives(iconsFolder, new[] { "Icon View.png", "Tile View.png", "LargeIcons.png", "large_icons.png" }, "LargeIcons") 
                    ?? GetBestShellIcon(new int[] { 269, 3, 2 }, "LargeIcons", true);
                    
                btnViewSmallIcons.Image = LoadWindowsXPIconWithAlternatives(iconsFolder, new[] { "Folder View - Classic.png", "Thumbnail View.png", "SmallIcons.png", "small_icons.png" }, "SmallIcons") 
                    ?? GetBestShellIcon(new int[] { 268, 152, 2 }, "SmallIcons", true);
                    
                btnViewList.Image = LoadWindowsXPIconWithAlternatives(iconsFolder, new[] { "Checklist.png", "List.png", "list.png", "List View.png" }, "List") 
                    ?? GetBestShellIcon(new int[] { 267, 151, 1 }, "List", true);
                    
                btnViewDetails.Image = LoadWindowsXPIconWithAlternatives(iconsFolder, new[] { "Detail View.png", "Details.png", "details.png", "Detail.png" }, "Details") 
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

        private Bitmap? LoadWindowsXPIconWithAlternatives(string iconsFolder, string[] iconFileNames, string buttonName)
        {
            foreach (string iconFileName in iconFileNames)
            {
                var bitmap = LoadWindowsXPIcon(iconsFolder, iconFileName);
                if (bitmap != null)
                {
                    Debug.WriteLine($"Successfully loaded {buttonName} icon from: {iconFileName}");
                    return bitmap;
                }
            }
            
            Debug.WriteLine($"Failed to load {buttonName} icon from any of the attempted filenames");
            return null;
        }

        private Bitmap? LoadWindowsXPIcon(string iconsFolder, string iconFileName)
        {
            try
            {
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

                // Load the image and resize to 16x16 for toolbar
                using (var originalImage = Image.FromFile(iconsPath))
                {
                    Debug.WriteLine($"Original image size: {originalImage.Width}x{originalImage.Height}, Format: {originalImage.PixelFormat}");
                    
                    var bitmap = new Bitmap(16, 16, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
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
                        g.DrawImage(originalImage, new Rectangle(0, 0, 16, 16), 
                            new Rectangle(0, 0, originalImage.Width, originalImage.Height), GraphicsUnit.Pixel);
                    }
                    
                    Debug.WriteLine($"Successfully created 16x16 bitmap from: {iconFileName}");
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
            // Enhanced fallback icons with better visual quality
            btnBack.Image = CreateEnhancedFallbackIcon("Back", 16);
            btnForward.Image = CreateEnhancedFallbackIcon("Forward", 16);
            btnUp.Image = CreateEnhancedFallbackIcon("Up", 16);
            btnRefresh.Image = CreateEnhancedFallbackIcon("Refresh", 16);
            
            // Fallback icons for view modes
            btnViewLargeIcons.Image = CreateViewIcon("LargeIcons", 16);
            btnViewSmallIcons.Image = CreateViewIcon("SmallIcons", 16);
            btnViewList.Image = CreateViewIcon("List", 16);
            btnViewDetails.Image = CreateViewIcon("Details", 16);
        }

        public Bitmap CreateViewIcon(string viewType, int size)
        {
            var bitmap = new Bitmap(size, size);
            using (var g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.Transparent);
                g.SmoothingMode = SmoothingMode.HighQuality;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                
                var blackBrush = Brushes.Black;
                var lightPen = new Pen(Color.FromArgb(255, 255, 255));
                var darkPen = new Pen(Color.FromArgb(128, 128, 128));
                
                switch (viewType)
                {
                    case "LargeIcons":
                        // Draw large icon grid (2x2)
                        g.FillRectangle(blackBrush, 2, 2, 4, 4);
                        g.FillRectangle(blackBrush, 8, 2, 4, 4);
                        g.FillRectangle(blackBrush, 2, 8, 4, 4);
                        g.FillRectangle(blackBrush, 8, 8, 4, 4);
                        break;
                        
                    case "SmallIcons":
                        // Draw small icon grid (3x3)
                        for (int i = 0; i < 3; i++)
                        {
                            for (int j = 0; j < 3; j++)
                            {
                                g.FillRectangle(blackBrush, 2 + i * 4, 2 + j * 4, 2, 2);
                            }
                        }
                        break;
                        
                    case "List":
                        // Draw list lines
                        for (int i = 0; i < 4; i++)
                        {
                            g.DrawLine(Pens.Black, 2, 3 + i * 3, size - 3, 3 + i * 3);
                        }
                        break;
                        
                    case "Details":
                        // Draw details view (lines with columns)
                        g.DrawLine(Pens.Black, 2, 2, size - 3, 2);
                        g.DrawLine(Pens.Black, 2, 2, 2, size - 3);
                        g.DrawLine(Pens.Black, size - 3, 2, size - 3, size - 3);
                        g.DrawLine(Pens.Black, 2, size - 3, size - 3, size - 3);
                        
                        // Vertical separators for columns
                        g.DrawLine(Pens.Black, size / 2, 2, size / 2, size - 3);
                        
                        // Horizontal lines for rows
                        for (int i = 1; i < 3; i++)
                        {
                            int y = 2 + (size - 5) * i / 3;
                            g.DrawLine(Pens.Black, 2, y, size - 3, y);
                        }
                        break;
                        
                    case "folders":
                        // Draw folder tree icon for TreeView toggle
                        // Draw a simple folder with tree structure
                        using (var folderBrush = new SolidBrush(Color.FromArgb(255, 255, 192)))
                        {
                            // Main folder
                            g.FillRectangle(folderBrush, 6, 8, 8, 6);
                            g.FillRectangle(folderBrush, 6, 6, 4, 3);
                        }
                        
                        // Folder outline
                        g.DrawRectangle(Pens.Black, 6, 8, 8, 6);
                        g.DrawRectangle(Pens.Black, 6, 6, 4, 3);
                        
                        // Tree lines to indicate folder hierarchy
                        g.DrawLine(Pens.Black, 2, 4, 2, 12);   // Vertical line
                        g.DrawLine(Pens.Black, 2, 9, 6, 9);    // Horizontal line to folder
                        g.DrawLine(Pens.Black, 2, 6, 4, 6);    // Branch line
                        g.FillRectangle(blackBrush, 1, 3, 2, 2);  // Small folder at top
                        break;
                        
                    case "star":
                        // Draw a classic Windows 95/98 style star for favorites/bookmarks
                        Point[] starPoints = {
                            new Point(size/2, 2),                    // Top point
                            new Point(size/2 + 2, size/2 - 1),      // Upper right
                            new Point(size - 2, size/2 - 1),        // Right point
                            new Point(size/2 + 3, size/2 + 2),      // Lower right
                            new Point(size/2 + 4, size - 2),        // Bottom right
                            new Point(size/2, size/2 + 4),          // Bottom point
                            new Point(size/2 - 4, size - 2),        // Bottom left
                            new Point(size/2 - 3, size/2 + 2),      // Lower left
                            new Point(2, size/2 - 1),               // Left point
                            new Point(size/2 - 2, size/2 - 1)       // Upper left
                        };
                        
                        // Fill the star with yellow color for classic appearance
                        using (var starBrush = new SolidBrush(Color.FromArgb(255, 255, 192)))
                        {
                            g.FillPolygon(starBrush, starPoints);
                        }
                        
                        // Draw star outline in black
                        g.DrawPolygon(Pens.Black, starPoints);
                        break;
                        
                    case "navigate":
                        // Draw a navigate/expand icon (folder with arrow)
                        // Draw folder background
                        using (var folderBrush = new SolidBrush(Color.FromArgb(255, 255, 192)))
                        {
                            g.FillRectangle(folderBrush, 2, 6, 10, 8);
                            g.FillRectangle(folderBrush, 2, 4, 4, 3);
                        }
                        
                        // Folder outline
                        g.DrawRectangle(Pens.Black, 2, 6, 10, 8);
                        g.DrawRectangle(Pens.Black, 2, 4, 4, 3);
                        
                        // Draw arrow pointing to tree
                        Point[] arrowPoints = {
                            new Point(size - 2, size/2 - 2),
                            new Point(size - 6, size/2),
                            new Point(size - 2, size/2 + 2)
                        };
                        g.FillPolygon(Brushes.Black, arrowPoints);
                        
                        // Draw simple tree lines
                        g.DrawLine(Pens.Black, size - 10, size/2 - 4, size - 10, size/2 + 4);
                        g.DrawLine(Pens.Black, size - 10, size/2, size - 7, size/2);
                        break;
                }
                
                lightPen.Dispose();
                darkPen.Dispose();
            }
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
                _ => CreateEnhancedFallbackIcon(buttonName, smallIcon ? 16 : 32)
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

        private Bitmap CreateEnhancedFallbackIcon(string iconType, int size)
        {
            var bitmap = new Bitmap(size, size);
            using (var g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.Transparent);
                g.SmoothingMode = SmoothingMode.HighQuality;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                
                // Use classic Windows 95/98 3D button styling
                var lightPen = new Pen(Color.FromArgb(255, 255, 255)); // Highlight
                var darkPen = new Pen(Color.FromArgb(128, 128, 128));   // Shadow
                var blackPen = new Pen(Color.Black);
                
                switch (iconType)
                {
                    case "Back":
                        // Draw classic 3D back arrow with Windows 95 styling
                        var backPoints = new Point[] 
                        { 
                            new(size-6, size/2-3), 
                            new(4, size/2), 
                            new(size-6, size/2+3) 
                        };
                        g.FillPolygon(Brushes.Black, backPoints);
                        // Add 3D highlight effect
                        g.DrawLine(lightPen, 4, size/2-1, size-7, size/2-4);
                        break;
                        
                    case "Forward":
                        // Draw classic 3D forward arrow
                        var forwardPoints = new Point[] 
                        { 
                            new(6, size/2-3), 
                            new(size-4, size/2), 
                            new(6, size/2+3) 
                        };
                        g.FillPolygon(Brushes.Black, forwardPoints);
                        // Add 3D highlight effect
                        g.DrawLine(lightPen, 7, size/2-4, size-4, size/2-1);
                        break;
                        
                    case "Up":
                        // Draw classic 3D up arrow
                        var upPoints = new Point[] 
                        { 
                            new(size/2, 4), 
                            new(size/2-4, size-6), 
                            new(size/2+4, size-6) 
                        };
                        g.FillPolygon(Brushes.Black, upPoints);
                        // Add 3D highlight effect
                        g.DrawLine(lightPen, size/2-3, size-7, size/2, 5);
                        break;
                        
                    case "Refresh":
                        // Draw classic refresh icon with 3D styling
                        g.DrawArc(blackPen, 3, 3, size-6, size-6, 45, 270);
                        // Draw arrow head
                        var refreshArrow = new Point[] 
                        { 
                            new(size-3, 3), 
                            new(size-6, 2), 
                            new(size-6, 6) 
                        };
                        g.FillPolygon(Brushes.Black, refreshArrow);
                        // Add 3D highlight
                        g.DrawArc(lightPen, 4, 4, size-8, size-8, 225, 90);
                        break;
                }
                
                // Clean up pens
                lightPen.Dispose();
                darkPen.Dispose();
                blackPen.Dispose();
            }
            return bitmap;
        }

        public Icon CreateApplicationIcon()
        {
            var bitmap = new Bitmap(32, 32);
            using (var g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.Transparent);
                
                // Draw a simple folder icon for the application
                using (var brush = new SolidBrush(Color.FromArgb(255, 255, 192)))
                {
                    g.FillRectangle(brush, 4, 12, 24, 16);
                    g.FillRectangle(brush, 4, 8, 10, 6);
                }
                
                using (var pen = new Pen(Color.Black, 2))
                {
                    g.DrawRectangle(pen, 4, 12, 24, 16);
                    g.DrawRectangle(pen, 4, 8, 10, 6);
                }
            }
            
            return Icon.FromHandle(bitmap.GetHicon());
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