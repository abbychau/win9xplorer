using System.Diagnostics;

namespace win9xplorer
{
    /// <summary>
    /// Manages navigation history and directory operations for the file explorer
    /// </summary>
    internal class NavigationManager
    {
        private string currentPath = "";
        private List<string> navigationHistory = new List<string>();
        private int currentHistoryIndex = -1;

        public string CurrentPath
        {
            get => currentPath;
            set => currentPath = value;
        }

        public bool CanGoBack => currentHistoryIndex > 0;
        public bool CanGoForward => currentHistoryIndex < navigationHistory.Count - 1;
        public bool CanGoUp => ExplorerUtils.GetParentPath(currentPath) != null;

        public void AddToHistory(string path)
        {
            // Update navigation history
            if (currentHistoryIndex == -1 || navigationHistory[currentHistoryIndex] != path)
            {
                // Remove forward history when navigating to new path
                if (currentHistoryIndex < navigationHistory.Count - 1)
                {
                    navigationHistory.RemoveRange(currentHistoryIndex + 1, 
                        navigationHistory.Count - currentHistoryIndex - 1);
                }
                
                navigationHistory.Add(path);
                currentHistoryIndex = navigationHistory.Count - 1;
            }
        }

        public string? GoBack()
        {
            if (CanGoBack)
            {
                currentHistoryIndex--;
                return navigationHistory[currentHistoryIndex];
            }
            return null;
        }

        public string? GoForward()
        {
            if (CanGoForward)
            {
                currentHistoryIndex++;
                return navigationHistory[currentHistoryIndex];
            }
            return null;
        }

        public string? GoUp()
        {
            return ExplorerUtils.GetParentPath(currentPath);
        }

        public void LoadDrives(TreeView treeView, ImageList imageListSmall, ImageList imageListLarge, IconManager iconManager)
        {
            treeView.Nodes.Clear();
            
            // Add "My Computer" root node
            var computerNode = new TreeNode("My Computer")
            {
                ImageKey = "computer",
                SelectedImageKey = "computer",
                Tag = "computer"
            };
            treeView.Nodes.Add(computerNode);

            // Add "Favorites" root node  
            var favoritesNode = new TreeNode("Favorites")
            {
                ImageKey = "favorites",
                SelectedImageKey = "favorites",
                Tag = "favorites"
            };
            
            // Add favorites icon if not exists
            if (!imageListSmall.Images.ContainsKey("favorites"))
            {
                // Create a simple star icon for favorites
                var starBitmap = iconManager.CreateViewIcon("star", 16);
                if (starBitmap != null)
                {
                    imageListSmall.Images.Add("favorites", starBitmap);
                    imageListLarge.Images.Add("favorites", iconManager.CreateViewIcon("star", 32) ?? starBitmap);
                }
            }
            
            treeView.Nodes.Add(favoritesNode);

            // Add drives to My Computer
            foreach (var drive in DriveInfo.GetDrives())
            {
                if (drive.IsReady)
                {
                    var driveInfo = ExplorerUtils.GetDriveInfo(drive);
                    
                    // Get system icon for this specific drive
                    string driveIconKey = $"drive_{drive.Name.Replace("\\", "")}";
                    if (!imageListSmall.Images.ContainsKey(driveIconKey))
                    {
                        Icon? smallDriveIcon = iconManager.GetSystemIcon(drive.Name, true, false);
                        Icon? largeDriveIcon = iconManager.GetSystemIcon(drive.Name, false, false);
                        
                        if (smallDriveIcon != null)
                            imageListSmall.Images.Add(driveIconKey, smallDriveIcon);
                        if (largeDriveIcon != null)
                            imageListLarge.Images.Add(driveIconKey, largeDriveIcon);
                    }
                    
                    var driveNode = new TreeNode(driveInfo.label)
                    {
                        ImageKey = driveIconKey,
                        SelectedImageKey = driveIconKey,
                        Tag = drive.Name
                    };
                    
                    // Add dummy node to show expand button
                    driveNode.Nodes.Add(new TreeNode("Loading..."));
                    computerNode.Nodes.Add(driveNode);
                }
            }

            computerNode.Expand();
            favoritesNode.Expand();
        }

        public void PopulateFavoritesNode(TreeView treeView, List<BookmarkManager.Bookmark> bookmarks, IconManager iconManager)
        {
            // Find the Favorites node
            TreeNode? favoritesNode = null;
            foreach (TreeNode node in treeView.Nodes)
            {
                if (node.Tag?.ToString() == "favorites")
                {
                    favoritesNode = node;
                    break;
                }
            }
            
            if (favoritesNode == null) return;
            
            // Clear existing bookmark nodes
            favoritesNode.Nodes.Clear();
            
            // Add bookmark nodes
            foreach (var bookmark in bookmarks.OrderBy(b => b.Name))
            {
                var bookmarkNode = new TreeNode(bookmark.Name)
                {
                    Tag = $"bookmark:{bookmark.Path}",
                    ToolTipText = bookmark.Path
                };
                
                // Set appropriate icon based on bookmark type
                if (bookmark.Path == "My Computer")
                {
                    bookmarkNode.ImageKey = "computer";
                    bookmarkNode.SelectedImageKey = "computer";
                }
                else if (Directory.Exists(bookmark.Path))
                {
                    // Get folder icon
                    bookmarkNode.ImageKey = "folder";
                    bookmarkNode.SelectedImageKey = "folder";
                }
                else
                {
                    // Default bookmark icon
                    bookmarkNode.ImageKey = "favorites";
                    bookmarkNode.SelectedImageKey = "favorites";
                }
                
                favoritesNode.Nodes.Add(bookmarkNode);
            }
        }

        public void SyncTreeViewSelection(TreeView treeView, string path)
        {
            // Find and select the corresponding tree node
            foreach (TreeNode computerNode in treeView.Nodes)
            {
                if (computerNode.Tag?.ToString() == "computer")
                {
                    foreach (TreeNode driveNode in computerNode.Nodes)
                    {
                        if (path.StartsWith(driveNode.Tag?.ToString() ?? ""))
                        {
                            SelectNodeByPath(treeView, driveNode, path);
                            break;
                        }
                    }
                }
            }
        }

        private void SelectNodeByPath(TreeView treeView, TreeNode node, string targetPath)
        {
            if (node.Tag?.ToString() == targetPath)
            {
                treeView.SelectedNode = node;
                return;
            }

            foreach (TreeNode childNode in node.Nodes)
            {
                if (targetPath.StartsWith(childNode.Tag?.ToString() ?? ""))
                {
                    if (!node.IsExpanded)
                        node.Expand();
                    SelectNodeByPath(treeView, childNode, targetPath);
                    break;
                }
            }
        }
    }
}