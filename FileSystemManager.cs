using System.Diagnostics;

namespace win9xplorer
{
    /// <summary>
    /// Manages file system operations like loading directory contents and file information
    /// </summary>
    internal class FileSystemManager
    {
        private readonly IconManager iconManager;

        public FileSystemManager(IconManager iconManager)
        {
            this.iconManager = iconManager;
        }

        public void LoadDirectoryContents(string path, ListView listView)
        {
            listView.Items.Clear();

            try
            {
                var items = new List<ListViewItem>();
                
                // Add directories first
                foreach (var directory in Directory.GetDirectories(path))
                {
                    var dirInfo = new DirectoryInfo(directory);
                    var item = new ListViewItem(dirInfo.Name)
                    {
                        ImageKey = iconManager.GetIconKey(directory, true),
                        Tag = directory
                    };
                    
                    // Always add subitems for details view
                    item.SubItems.Add(""); // Size - empty for folders
                    item.SubItems.Add("File Folder");
                    item.SubItems.Add(dirInfo.LastWriteTime.ToString("M/d/yyyy h:mm tt"));
                    
                    items.Add(item);
                }

                // Add files
                foreach (var file in Directory.GetFiles(path))
                {
                    var fileInfo = new FileInfo(file);
                    var item = new ListViewItem(fileInfo.Name)
                    {
                        ImageKey = iconManager.GetIconKey(file, false),
                        Tag = file
                    };
                    
                    // Always add subitems for details view
                    item.SubItems.Add(ExplorerUtils.FormatFileSize(fileInfo.Length));
                    item.SubItems.Add(ExplorerUtils.GetFileType(fileInfo.Extension));
                    item.SubItems.Add(fileInfo.LastWriteTime.ToString("M/d/yyyy h:mm tt"));
                    
                    items.Add(item);
                }
                
                listView.Items.AddRange(items.ToArray());
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading directory contents: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        public void LoadMyComputerContents(ListView listView, ImageList imageListSmall, ImageList imageListLarge)
        {
            try
            {
                listView.Items.Clear();

                var items = new List<ListViewItem>();
                
                // Add all available drives
                foreach (var drive in DriveInfo.GetDrives())
                {
                    try
                    {
                        var driveInfo = ExplorerUtils.GetDriveInfo(drive);

                        // Use the same drive icon key as in the tree view
                        string driveIconKey = $"drive_{drive.Name.Replace("\\", "")}";
                        
                        // Ensure the drive icon exists in the image lists
                        if (!imageListSmall.Images.ContainsKey(driveIconKey))
                        {
                            Icon? smallDriveIcon = iconManager.GetSystemIcon(drive.Name, true, false);
                            Icon? largeDriveIcon = iconManager.GetSystemIcon(drive.Name, false, false);
                            
                            if (smallDriveIcon != null)
                                imageListSmall.Images.Add(driveIconKey, smallDriveIcon);
                            if (largeDriveIcon != null)
                                imageListLarge.Images.Add(driveIconKey, largeDriveIcon);
                        }
                        
                        var item = new ListViewItem(driveInfo.label)
                        {
                            ImageKey = driveIconKey,
                            Tag = drive.Name
                        };
                        
                        // Add subitems for My Computer details view
                        if (listView.View == View.Details)
                        {
                            item.SubItems.Add(driveInfo.totalSize);      // Total Size
                            item.SubItems.Add(driveInfo.freeSpace);      // Free Space  
                            item.SubItems.Add(driveInfo.usedSpace);      // Used Space
                            item.SubItems.Add($"{drive.DriveType} Drive"); // Type
                            item.SubItems.Add(driveInfo.fileSystem);     // File System
                        }
                        else
                        {
                            // For other views, add the standard subitems for compatibility
                            item.SubItems.Add(driveInfo.totalSize);      // Size (shown in other views too)
                            item.SubItems.Add($"{drive.DriveType} Drive"); // Type
                            item.SubItems.Add("");             // Modified (empty for drives)
                        }
                        
                        items.Add(item);
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Error processing drive {drive.Name}: {ex.Message}");
                        // Continue with other drives even if one fails
                    }
                }
                
                listView.Items.AddRange(items.ToArray());
                
                Debug.WriteLine($"Loaded My Computer view with {items.Count} drives");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading My Computer contents: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Debug.WriteLine($"Error in LoadMyComputerContents: {ex.Message}");
            }
        }

        public void SetupDetailsView(ListView listView)
        {
            listView.Columns.Clear();
            listView.Columns.Add("Name", 200);
            listView.Columns.Add("Size", 100);
            listView.Columns.Add("Type", 120);
            listView.Columns.Add("Modified", 150);
        }

        public void SetupMyComputerDetailsView(ListView listView)
        {
            listView.Columns.Clear();
            listView.Columns.Add("Name", 250);           // Drive name and label
            listView.Columns.Add("Total Size", 100);     // Total drive capacity
            listView.Columns.Add("Free Space", 100);     // Available free space
            listView.Columns.Add("Used Space", 100);     // Used space
            listView.Columns.Add("Type", 120);           // Drive type (Fixed, Removable, etc.)
            listView.Columns.Add("File System", 80);     // NTFS, FAT32, etc.
        }

        public void ExpandTreeNode(TreeNode node)
        {
            // Clear dummy nodes and load actual folders
            node.Nodes.Clear();
            
            try
            {
                string path = node.Tag?.ToString() ?? "";
                if (Directory.Exists(path))
                {
                    foreach (var directory in Directory.GetDirectories(path))
                    {
                        var folderNode = new TreeNode(Path.GetFileName(directory))
                        {
                            ImageKey = "folder",
                            SelectedImageKey = "folder",
                            Tag = directory
                        };

                        // Check if folder has subdirectories
                        try
                        {
                            if (Directory.GetDirectories(directory).Length > 0)
                            {
                                folderNode.Nodes.Add(new TreeNode("Loading..."));
                            }
                        }
                        catch { }

                        node.Nodes.Add(folderNode);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading folders: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
}