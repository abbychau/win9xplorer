using System.Diagnostics;
using System.Runtime.InteropServices;

namespace win9xplorer
{
    /// <summary>
    /// Manages context menus for the file explorer
    /// </summary>
    internal class ContextMenuManager
    {
        private ContextMenuStrip? listViewContextMenu;
        private Form? parentForm;
        private ContextMenuStrip? currentActiveMenu;
        private Action? renameCallback;
        private Action? deleteCallback;

        public void SetupContextMenu(ListView listView)
        {
            // Store reference to parent form
            parentForm = listView.FindForm();
            
            listViewContextMenu = new ContextMenuStrip();
            
            var refreshItem = new ToolStripMenuItem("Refresh");
            
            var viewMenu = new ToolStripMenuItem("View");
            var largeIconsItem = new ToolStripMenuItem("Large Icons");
            var smallIconsItem = new ToolStripMenuItem("Small Icons");
            var listItem = new ToolStripMenuItem("List");
            var detailsItem = new ToolStripMenuItem("Details");
            
            viewMenu.DropDownItems.AddRange(new[] { largeIconsItem, smallIconsItem, listItem, detailsItem });
            
            listViewContextMenu.Items.AddRange(new ToolStripItem[] {
                refreshItem,
                new ToolStripSeparator(),
                viewMenu
            });
        }

        public void SetupRenameCallback(Action renameCallback)
        {
            this.renameCallback = renameCallback;
        }

        public void SetupDeleteCallback(Action deleteCallback)
        {
            this.deleteCallback = deleteCallback;
        }

        public void SetupTreeViewContextMenu(TreeView treeView)
        {
            // Store reference to parent form if not already set
            if (parentForm == null)
                parentForm = treeView.FindForm();
                
            // Use MouseUp instead of MouseDown for proper context menu behavior
            treeView.MouseUp += TreeView_MouseUp;
            // Also add MouseDown to cancel existing menus
            treeView.MouseDown += TreeView_MouseDown;
        }

        private void TreeView_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                Debug.WriteLine("TreeView right-click down - cancelling any existing context menu");
                CancelCurrentContextMenu();
            }
        }

        private void TreeView_MouseUp(object sender, MouseEventArgs e)
        {
            if (sender is not TreeView treeView || e.Button != MouseButtons.Right)
                return;
                
            Debug.WriteLine($"TreeView right-click up at: {e.Location}");
                
            // Get the node at the click position
            var node = treeView.GetNodeAt(e.Location);
            if (node != null)
            {
                // Select the node that was right-clicked
                treeView.SelectedNode = node;
                
                var screenPoint = treeView.PointToScreen(e.Location);
                string nodeTag = node.Tag?.ToString() ?? "";
                
                Debug.WriteLine($"Right-clicked on tree node: {nodeTag}");
                
                if (nodeTag == "computer" || nodeTag == "favorites")
                {
                    // Show simple context menu for special nodes
                    ShowSpecialNodeContextMenu(nodeTag, screenPoint.X, screenPoint.Y);
                }
                else if (nodeTag.StartsWith("bookmark:"))
                {
                    // Handle bookmark nodes
                    string bookmarkPath = nodeTag.Substring("bookmark:".Length);
                    if (Directory.Exists(bookmarkPath) || File.Exists(bookmarkPath))
                    {
                        ShowSystemContextMenu(bookmarkPath, screenPoint.X, screenPoint.Y);
                    }
                }
                else if (Directory.Exists(nodeTag))
                {
                    // Regular folder/drive - show OS context menu
                    ShowSystemContextMenu(nodeTag, screenPoint.X, screenPoint.Y);
                }
            }
        }

        private void ShowSpecialNodeContextMenu(string nodeType, int x, int y)
        {
            // Cancel any existing context menu first
            CancelCurrentContextMenu();
            
            var contextMenu = new ContextMenuStrip();
            
            // Track this as the current active menu
            currentActiveMenu = contextMenu;
            
            if (nodeType == "computer")
            {
                var refreshItem = new ToolStripMenuItem("Refresh");
                refreshItem.Click += (s, e) => {
                    MessageBox.Show("Computer refresh functionality would be implemented here.", "Refresh", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                };
                
                contextMenu.Items.Add(refreshItem);
            }
            else if (nodeType == "favorites")
            {
                var organizeItem = new ToolStripMenuItem("Organize Favorites");
                organizeItem.Click += (s, e) => {
                    MessageBox.Show("Organize Favorites functionality would be implemented here.", "Organize Favorites", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                };
                
                contextMenu.Items.Add(organizeItem);
            }
            
            // Handle menu closed event to clear tracking
            contextMenu.Closed += (s, e) => {
                if (currentActiveMenu == contextMenu)
                    currentActiveMenu = null;
            };
            
            contextMenu.Show(x, y);
        }

        public void ShowContextMenu(ListView listView, MouseEventArgs e, string currentPath, 
            Action refreshAction, Action<View> setViewAction)
        {
            // This method should now be called from MouseUp, not MouseDown
            if (e.Button == MouseButtons.Right)
            {
                Debug.WriteLine($"Right-click up detected at: {e.Location}, CurrentPath: '{currentPath}'");
                
                // Get the item at the mouse position
                var hitTest = listView.HitTest(e.Location);
                
                if (hitTest.Item != null)
                {
                    // Check if the right-clicked item is already part of the selection
                    if (!hitTest.Item.Selected)
                    {
                        // If not selected, clear selection and select only this item
                        listView.SelectedItems.Clear();
                        hitTest.Item.Selected = true;
                    }
                    // If it's already selected, keep the existing selection (multiple files)
                    
                    // Get all selected items for context menu
                    var selectedPaths = new List<string>();
                    foreach (ListViewItem item in listView.SelectedItems)
                    {
                        string itemPath = item.Tag?.ToString() ?? "";
                        if (!string.IsNullOrEmpty(itemPath))
                        {
                            selectedPaths.Add(itemPath);
                        }
                    }
                    
                    Debug.WriteLine($"Right-clicked on items (count: {selectedPaths.Count}): {string.Join(", ", selectedPaths.Take(3))}{(selectedPaths.Count > 3 ? "..." : "")}");
                    
                    if (selectedPaths.Count > 0)
                    {
                        // Convert ListView coordinates to screen coordinates
                        var screenPoint = listView.PointToScreen(e.Location);
                        
                        // Special handling for My Computer view (drives)
                        if (string.IsNullOrEmpty(currentPath))
                        {
                            // This is a drive in My Computer view
                            Debug.WriteLine("Showing drive context menu");
                            if (selectedPaths.Count == 1)
                            {
                                ShowSystemContextMenu(selectedPaths[0], screenPoint.X, screenPoint.Y);
                            }
                            else
                            {
                                // For multiple drives, show native Windows context menu
                                ShowNativeMultipleItemsContextMenu(selectedPaths, screenPoint.X, screenPoint.Y);
                            }
                        }
                        else
                        {
                            // Regular files or folders - always show native Windows context menu
                            Debug.WriteLine("Showing file/folder context menu");
                            if (selectedPaths.Count == 1)
                            {
                                ShowSystemContextMenu(selectedPaths[0], screenPoint.X, screenPoint.Y);
                            }
                            else
                            {
                                // For multiple files/folders, show native Windows context menu
                                ShowNativeMultipleItemsContextMenu(selectedPaths, screenPoint.X, screenPoint.Y);
                            }
                        }
                    }
                }
                else
                {
                    // Right-clicked on empty space - show folder context menu or custom menu
                    var screenPoint = listView.PointToScreen(e.Location);
                    Debug.WriteLine($"Right-clicked on empty space, screen point: {screenPoint}");
                    
                    if (!string.IsNullOrEmpty(currentPath) && Directory.Exists(currentPath))
                    {
                        // Show OS context menu for the current folder
                        Debug.WriteLine("Showing OS folder context menu");
                        ShowOSContextMenuForFolder(currentPath, screenPoint.X, screenPoint.Y);
                    }
                    else
                    {
                        // Show our custom context menu for My Computer view
                        Debug.WriteLine("Showing custom context menu for special view");
                        UpdateContextMenuEventHandlers(refreshAction, setViewAction);
                        
                        // Track this as the current active menu
                        currentActiveMenu = listViewContextMenu;
                        
                        // Handle menu closed event to clear tracking
                        if (listViewContextMenu != null)
                        {
                            listViewContextMenu.Closed += (s, args) => {
                                if (currentActiveMenu == listViewContextMenu)
                                    currentActiveMenu = null;
                            };
                            
                            listViewContextMenu.Show(screenPoint);
                        }
                    }
                }
            }
        }

        private void UpdateContextMenuEventHandlers(Action refreshAction, Action<View> setViewAction)
        {
            if (listViewContextMenu == null) return;

            // Clear existing handlers and set new ones
            foreach (ToolStripItem item in listViewContextMenu.Items)
            {
                if (item is ToolStripMenuItem menuItem)
                {
                    if (menuItem.Text == "Refresh")
                    {
                        // Remove all existing click handlers
                        var field = typeof(ToolStripMenuItem).GetField("EventClick", 
                            System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic);
                        if (field?.GetValue(menuItem) != null)
                        {
                            menuItem.Click -= (s, e) => refreshAction();
                        }
                        
                        menuItem.Click += (s, e) => refreshAction();
                    }
                    else if (menuItem.Text == "View" && menuItem.HasDropDownItems)
                    {
                        foreach (ToolStripMenuItem viewItem in menuItem.DropDownItems.OfType<ToolStripMenuItem>())
                        {
                            // Clear and reassign view handlers
                            viewItem.Click -= (s, e) => { };
                            
                            switch (viewItem.Text)
                            {
                                case "Large Icons":
                                    viewItem.Click += (s, e) => setViewAction(View.LargeIcon);
                                    break;
                                case "Small Icons":
                                    viewItem.Click += (s, e) => setViewAction(View.SmallIcon);
                                    break;
                                case "List":
                                    viewItem.Click += (s, e) => setViewAction(View.List);
                                    break;
                                case "Details":
                                    viewItem.Click += (s, e) => setViewAction(View.Details);
                                    break;
                            }
                        }
                    }
                }
            }
        }

        private void ShowWindowsContextMenuForDrive(ListView listView, string drivePath, Point location)
        {
            try
            {
                // For drives, we need to handle them specially
                // Convert ListView coordinates to screen coordinates
                var screenPoint = listView.PointToScreen(location);
                
                // Show the system context menu for the drive
                ShowSystemContextMenu(drivePath, screenPoint.X, screenPoint.Y);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error showing Windows context menu for drive: {ex.Message}");
                // Fallback to a simple message if the system context menu fails
                MessageBox.Show($"Right-clicked on drive: {drivePath}", "Drive Context Menu", 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void ShowWindowsContextMenu(ListView listView, string filePath, Point location)
        {
            try
            {
                // Convert ListView coordinates to screen coordinates
                var screenPoint = listView.PointToScreen(location);
                
                // Show the system context menu for the file/folder
                ShowSystemContextMenu(filePath, screenPoint.X, screenPoint.Y);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error showing Windows context menu: {ex.Message}");
                // Fallback to a simple message if the system context menu fails
                MessageBox.Show($"Right-clicked on: {Path.GetFileName(filePath)}", "Context Menu", 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void ShowSystemContextMenu(string filePath, int x, int y)
        {
            Debug.WriteLine($"ShowSystemContextMenu called for: {filePath} at ({x}, {y})");
            
            // Cancel any existing context menu first
            CancelCurrentContextMenu();
            
            try
            {
                // Try the COM interface approach first
                if (TryShowComContextMenu(filePath, x, y))
                {
                    Debug.WriteLine("COM context menu succeeded");
                    return;
                }

                Debug.WriteLine("COM context menu failed, trying simple fallback");
                // Fallback: Use a simpler approach
                ShowSimpleContextMenu(filePath, x, y);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Exception in ShowSystemContextMenu: {ex.Message}");
                ShowSimpleContextMenu(filePath, x, y);
            }
        }

        private bool TryShowComContextMenu(string filePath, int x, int y)
        {
            Debug.WriteLine($"TryShowComContextMenu for: {filePath}");
            
            try
            {
                IntPtr pidl = IntPtr.Zero;
                IntPtr parentFolder = IntPtr.Zero;
                IntPtr contextMenu = IntPtr.Zero;
                IntPtr menu = IntPtr.Zero;

                try
                {
                    // Get the PIDL for the file
                    uint sfgao = 0;
                    int result = WinApi.SHParseDisplayName(filePath, IntPtr.Zero, out pidl, 0, out sfgao);
                    if (result != 0 || pidl == IntPtr.Zero)
                    {
                        Debug.WriteLine($"Failed to get PIDL for {filePath}, HRESULT: 0x{result:X}");
                        return false;
                    }

                    // Get the parent folder
                    IntPtr pidlLast = IntPtr.Zero;
                    result = WinApi.SHBindToParent(pidl, WinApi.IID_IShellFolder, out parentFolder, out pidlLast);
                    if (result != 0 || parentFolder == IntPtr.Zero)
                    {
                        Debug.WriteLine($"Failed to get parent folder for {filePath}, HRESULT: 0x{result:X}");
                        return false;
                    }

                    // Get the IShellFolder interface
                    var shellFolder = Marshal.GetObjectForIUnknown(parentFolder) as WinApi.IShellFolder;
                    if (shellFolder == null)
                    {
                        Debug.WriteLine("Failed to get IShellFolder interface");
                        return false;
                    }

                    // Get the context menu
                    IntPtr[] pidlArray = { pidlLast };
                    var contextMenuGuid = WinApi.IID_IContextMenu;
                    
                    // Get parent form handle for proper context menu positioning
                    IntPtr parentHwnd = parentForm?.Handle ?? IntPtr.Zero;
                    
                    result = shellFolder.GetUIObjectOf(parentHwnd, 1, pidlArray, ref contextMenuGuid, IntPtr.Zero, out contextMenu);
                    if (result != 0 || contextMenu == IntPtr.Zero)
                    {
                        Debug.WriteLine($"Failed to get context menu interface, HRESULT: 0x{result:X}");
                        return false;
                    }

                    var contextMenuInterface = Marshal.GetObjectForIUnknown(contextMenu) as WinApi.IContextMenu;
                    if (contextMenuInterface == null)
                    {
                        Debug.WriteLine("Failed to get IContextMenu interface");
                        return false;
                    }

                    // Create a popup menu
                    menu = WinApi.CreatePopupMenu();
                    if (menu == IntPtr.Zero)
                    {
                        Debug.WriteLine("Failed to create popup menu");
                        return false;
                    }

                    // Let the shell add its menu items
                    result = contextMenuInterface.QueryContextMenu(menu, 0, 1, 0x7FFF, 0x10); // CMF_NORMAL
                    if (result < 0)
                    {
                        Debug.WriteLine($"Failed to query context menu, HRESULT: 0x{result:X}");
                        return false;
                    }

                    Debug.WriteLine($"Context menu populated successfully, showing at ({x}, {y})");

                    // Show the menu
                    uint command = (uint)WinApi.TrackPopupMenuEx(menu, WinApi.TPM_RETURNCMD | WinApi.TPM_LEFTBUTTON, x, y, parentHwnd, IntPtr.Zero);
                    if (command > 0)
                    {
                        Debug.WriteLine($"Context menu command selected: {command}");
                        
                        // Execute the selected command
                        var commandInfo = new WinApi.CMINVOKECOMMANDINFO
                        {
                            cbSize = Marshal.SizeOf(typeof(WinApi.CMINVOKECOMMANDINFO)),
                            fMask = 0,
                            hwnd = parentHwnd,
                            lpVerb = new IntPtr(command - 1),
                            lpParameters = IntPtr.Zero,
                            lpDirectory = IntPtr.Zero,
                            nShow = 1, // SW_SHOWNORMAL
                            dwHotKey = 0,
                            hIcon = IntPtr.Zero
                        };

                        IntPtr commandInfoPtr = Marshal.AllocHGlobal(Marshal.SizeOf(commandInfo));
                        try
                        {
                            Marshal.StructureToPtr(commandInfo, commandInfoPtr, false);
                            result = contextMenuInterface.InvokeCommand(commandInfoPtr);
                            Debug.WriteLine($"Command invocation result: 0x{result:X}");
                        }
                        finally
                        {
                            Marshal.FreeHGlobal(commandInfoPtr);
                        }
                    }
                    else
                    {
                        Debug.WriteLine("No command selected (user cancelled or clicked outside)");
                    }

                    return true;
                }
                finally
                {
                    // Clean up
                    if (menu != IntPtr.Zero)
                        WinApi.DestroyMenu(menu);
                    if (contextMenu != IntPtr.Zero)
                        Marshal.Release(contextMenu);
                    if (parentFolder != IntPtr.Zero)
                        Marshal.Release(parentFolder);
                    if (pidl != IntPtr.Zero)
                        WinApi.CoTaskMemFree(pidl);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Exception in TryShowComContextMenu: {ex.Message}");
                Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                return false;
            }
        }

        private void ShowSimpleContextMenu(string filePath, int x, int y)
        {
            Debug.WriteLine($"ShowSimpleContextMenu called for: {filePath} at ({x}, {y})");
            
            try
            {
                // Create a simple context menu with common actions
                var contextMenu = new ContextMenuStrip();
                
                // Track this as the current active menu
                currentActiveMenu = contextMenu;
                
                // Open
                var openItem = new ToolStripMenuItem("Open");
                openItem.Click += (s, e) => {
                    try
                    {
                        Debug.WriteLine($"Opening: {filePath}");
                        Process.Start(new ProcessStartInfo
                        {
                            FileName = filePath,
                            UseShellExecute = true
                        });
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error opening {Path.GetFileName(filePath)}: {ex.Message}", 
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                };
                
                // Open in Explorer (for folders) or Open file location (for files)
                var exploreItem = new ToolStripMenuItem(Directory.Exists(filePath) ? "Explore" : "Open file location");
                exploreItem.Click += (s, e) => {
                    try
                    {
                        if (Directory.Exists(filePath))
                        {
                            Process.Start("explorer.exe", filePath);
                        }
                        else
                        {
                            Process.Start("explorer.exe", $"/select,\"{filePath}\"");
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error opening explorer: {ex.Message}", 
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                };
                
                // Properties
                var propertiesItem = new ToolStripMenuItem("Properties");
                propertiesItem.Click += (s, e) => {
                    try
                    {
                        ShowFileProperties(filePath);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error showing properties: {ex.Message}", 
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                };
                
                // Cut, Copy, Paste
                var cutItem = new ToolStripMenuItem("Cut");
                cutItem.Click += (s, e) => {
                    Debug.WriteLine($"Cut invoked for: {filePath}");
                    CutFilesToClipboard(new List<string> { filePath });
                };
                
                var copyItem = new ToolStripMenuItem("Copy");
                copyItem.Click += (s, e) => {
                    Debug.WriteLine($"Copy invoked for: {filePath}");
                    CopyFilesToClipboard(new List<string> { filePath });
                };
                
                var pasteItem = new ToolStripMenuItem("Paste");
                pasteItem.Click += (s, e) => {
                    MessageBox.Show("Paste functionality - individual file context menu", "Paste", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                };
                
                // Rename
                var renameItem = new ToolStripMenuItem("Rename");
                renameItem.Click += (s, e) => {
                    renameCallback?.Invoke();
                };
                
                // Delete
                var deleteItem = new ToolStripMenuItem("Delete");
                deleteItem.Click += (s, e) => {
                    deleteCallback?.Invoke();
                };
                
                contextMenu.Items.AddRange(new ToolStripItem[] {
                    openItem,
                    exploreItem,
                    new ToolStripSeparator(),
                    propertiesItem,
                    new ToolStripSeparator(),
                    cutItem,
                    copyItem,
                    pasteItem,
                    new ToolStripSeparator(),
                    deleteItem,
                    renameItem
                });
                
                // Handle menu closed event to clear tracking
                contextMenu.Closed += (s, e) => {
                    if (currentActiveMenu == contextMenu)
                        currentActiveMenu = null;
                };
                
                Debug.WriteLine($"Showing simple context menu with {contextMenu.Items.Count} items");
                contextMenu.Show(x, y);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error in ShowSimpleContextMenu: {ex.Message}");
                // Last resort - show a message box at least
                MessageBox.Show($"Right-clicked on: {Path.GetFileName(filePath)}\n\nContext menu error: {ex.Message}", 
                    "Context Menu", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void ShowFileProperties(string filePath)
        {
            try
            {
                // Use the Shell API to show properties dialog
                var startInfo = new ProcessStartInfo
                {
                    FileName = "rundll32.exe",
                    Arguments = $"shell32.dll,OpenAs_RunDLL \"{filePath}\"",
                    UseShellExecute = true,
                    WindowStyle = ProcessWindowStyle.Hidden
                };
                
                // Alternative: show properties using shell32 Properties
                startInfo.Arguments = $"shell32.dll,SHObjectProperties 0,\"{filePath}\"";
                
                Process.Start(startInfo);
            }
            catch
            {
                // Alternative method: select the file in explorer (which often shows properties)
                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = "explorer.exe",
                        Arguments = $"/select,\"{filePath}\"",
                        UseShellExecute = true
                    });
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Unable to show properties for {Path.GetFileName(filePath)}: {ex.Message}", 
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private void ShowOSContextMenuForFolder(string folderPath, int x, int y)
        {
            try
            {
                // Show OS context menu for the folder itself
                ShowSystemContextMenu(folderPath, x, y);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error showing OS context menu for folder: {ex.Message}");
                // Fallback to simple context menu
                ShowSimpleFolderContextMenu(folderPath, x, y);
            }
        }

        private void ShowSimpleFolderContextMenu(string folderPath, int x, int y)
        {
            var contextMenu = new ContextMenuStrip();
            
            // Track this as the current active menu
            currentActiveMenu = contextMenu;
            
            var pasteItem = new ToolStripMenuItem("Paste");
            pasteItem.Click += (s, e) => {
                MessageBox.Show("Paste functionality would be implemented here.", "Paste", 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            };
            
            var newFolderItem = new ToolStripMenuItem("New Folder");
            newFolderItem.Click += (s, e) => {
                MessageBox.Show("New Folder functionality would be implemented here.", "New Folder", 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            };
            
            var propertiesItem = new ToolStripMenuItem("Properties");
            propertiesItem.Click += (s, e) => {
                try
                {
                    ShowFileProperties(folderPath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error showing properties: {ex.Message}", 
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            };
            
            contextMenu.Items.AddRange(new ToolStripItem[] {
                pasteItem,
                new ToolStripSeparator(),
                newFolderItem,
                new ToolStripSeparator(),
                propertiesItem
            });
            
            // Handle menu closed event to clear tracking
            contextMenu.Closed += (s, e) => {
                if (currentActiveMenu == contextMenu)
                    currentActiveMenu = null;
            };
            
            contextMenu.Show(x, y);
        }

        private void CancelCurrentContextMenu()
        {
            if (currentActiveMenu != null && currentActiveMenu.Visible)
            {
                Debug.WriteLine("Closing active context menu");
                currentActiveMenu.Close();
                currentActiveMenu = null;
            }
        }

        public void OnListViewMouseDown(MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                Debug.WriteLine("ListView right-click down - cancelling any existing context menu");
                CancelCurrentContextMenu();
            }
        }

        private void CopyFilesToClipboard(List<string> filePaths)
        {
            try
            {
                if (filePaths.Count == 0) return;
                
                // Create a StringCollection with file paths
                var files = new System.Collections.Specialized.StringCollection();
                files.AddRange(filePaths.ToArray());
                
                // Copy to clipboard
                Clipboard.SetFileDropList(files);
                
                Debug.WriteLine($"Copied {filePaths.Count} files to clipboard");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error copying files to clipboard: {ex.Message}");
                MessageBox.Show($"Error copying files: {ex.Message}", "Copy Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void CutFilesToClipboard(List<string> filePaths)
        {
            try
            {
                if (filePaths.Count == 0) return;
                
                // Create a StringCollection with file paths
                var files = new System.Collections.Specialized.StringCollection();
                files.AddRange(filePaths.ToArray());
                
                // Copy to clipboard with cut format
                var dataObject = new DataObject();
                dataObject.SetFileDropList(files);
                
                // Add a special format to indicate cut (not copy)
                byte[] moveEffect = BitConverter.GetBytes(2); // DROPEFFECT_MOVE
                dataObject.SetData("Preferred DropEffect", moveEffect);
                
                Clipboard.SetDataObject(dataObject, true);
                
                Debug.WriteLine($"Cut {filePaths.Count} files to clipboard");
                
                // Show visual feedback that files are cut (could implement dimming effect later)
                MessageBox.Show($"Cut {filePaths.Count} file(s) to clipboard", "Cut", 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error cutting files to clipboard: {ex.Message}");
                MessageBox.Show($"Error cutting files: {ex.Message}", "Cut Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void DeleteFiles(List<string> filePaths)
        {
            try
            {
                if (filePaths.Count == 0) return;
                
                string message = filePaths.Count == 1 
                    ? $"Are you sure you want to delete '{Path.GetFileName(filePaths[0])}'?"
                    : $"Are you sure you want to delete these {filePaths.Count} items?";
                
                var result = MessageBox.Show(message, "Delete Files", 
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                
                if (result == DialogResult.Yes)
                {
                    int successCount = 0;
                    var errors = new List<string>();
                    
                    foreach (string filePath in filePaths)
                    {
                        try
                        {
                            if (Directory.Exists(filePath))
                            {
                                Directory.Delete(filePath, true); // Recursive delete for directories
                                successCount++;
                                Debug.WriteLine($"Deleted directory: {filePath}");
                            }
                            else if (File.Exists(filePath))
                            {
                                File.Delete(filePath);
                                successCount++;
                                Debug.WriteLine($"Deleted file: {filePath}");
                            }
                        }
                        catch (UnauthorizedAccessException)
                        {
                            errors.Add($"'{Path.GetFileName(filePath)}': Access denied");
                        }
                        catch (DirectoryNotFoundException)
                        {
                            errors.Add($"'{Path.GetFileName(filePath)}': Path not found");
                        }
                        catch (IOException ex)
                        {
                            errors.Add($"'{Path.GetFileName(filePath)}': {ex.Message}");
                        }
                        catch (Exception ex)
                        {
                            errors.Add($"'{Path.GetFileName(filePath)}': {ex.Message}");
                        }
                    }
                    
                    Debug.WriteLine($"Delete operation completed: {successCount} successful, {errors.Count} errors");
                    
                    // Refresh using the Form1's delete callback which will handle refresh properly
                    try
                    {
                        if (parentForm is Form1 mainForm)
                        {
                            // Use BeginInvoke to ensure UI refresh happens after COM context menu closes
                            mainForm.BeginInvoke(new Action(() => {
                                // Trigger proper refresh of the current directory
                                mainForm.GetType().GetMethod("BtnRefresh_Click", 
                                    System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)
                                    ?.Invoke(mainForm, new object[] { mainForm, EventArgs.Empty });
                            }));
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Error refreshing view after delete: {ex.Message}");
                    }
                    
                    // Show results if there were errors
                    if (errors.Count > 0)
                    {
                        if (successCount > 0)
                        {
                            string errorMessage = $"Successfully deleted {successCount} items.\n\nErrors:\n" + 
                                                string.Join("\n", errors.Take(5));
                            if (errors.Count > 5)
                                errorMessage += $"\n... and {errors.Count - 5} more errors";
                                
                            MessageBox.Show(errorMessage, "Delete Partial Success", 
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                        else
                        {
                            string errorMessage = "Failed to delete items:\n" + 
                                                string.Join("\n", errors.Take(5));
                            if (errors.Count > 5)
                                errorMessage += $"\n... and {errors.Count - 5} more errors";
                                
                            MessageBox.Show(errorMessage, "Delete Error", 
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error in delete operation: {ex.Message}");
                MessageBox.Show($"Error deleting files: {ex.Message}", "Delete Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void ShowMultipleFileProperties(List<string> filePaths)
        {
            try
            {
                if (filePaths.Count == 0) return;
                
                if (filePaths.Count == 1)
                {
                    // Single file - show regular properties
                    ShowFileProperties(filePaths[0]);
                }
                else
                {
                    // Multiple files - show summary
                    long totalSize = 0;
                    int fileCount = 0;
                    int dirCount = 0;
                    
                    foreach (string path in filePaths)
                    {
                        try
                        {
                            if (Directory.Exists(path))
                            {
                                dirCount++;
                                // Calculate directory size (basic implementation)
                                var dirInfo = new DirectoryInfo(path);
                                totalSize += dirInfo.EnumerateFiles("*", SearchOption.AllDirectories)
                                    .Sum(fi => fi.Length);
                            }
                            else if (File.Exists(path))
                            {
                                fileCount++;
                                var fileInfo = new FileInfo(path);
                                totalSize += fileInfo.Length;
                            }
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"Error calculating size for {path}: {ex.Message}");
                        }
                    }
                    
                    string sizeText = totalSize < 1024 ? $"{totalSize} bytes" :
                                     totalSize < 1048576 ? $"{totalSize / 1024:N1} KB" :
                                     totalSize < 1073741824 ? $"{totalSize / 1048576:N1} MB" :
                                     $"{totalSize / 1073741824:N1} GB";
                    
                    string message = $"Selected items: {filePaths.Count}\n";
                    if (fileCount > 0) message += $"Files: {fileCount}\n";
                    if (dirCount > 0) message += $"Folders: {dirCount}\n";
                    message += $"Total size: {sizeText}";
                    
                    MessageBox.Show(message, "Properties - Multiple Items", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error showing multiple file properties: {ex.Message}");
                MessageBox.Show($"Error showing properties: {ex.Message}", "Properties Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void ShowMultipleItemsContextMenu(ListView listView, List<string> selectedPaths, Point location)
        {
            try
            {
                // Convert ListView coordinates to screen coordinates if listView is not null
                Point screenPoint;
                if (listView != null)
                {
                    screenPoint = listView.PointToScreen(location);
                }
                else
                {
                    screenPoint = location; // Assume location is already in screen coordinates
                }
                
                // For multiple items, show an enhanced custom context menu with comprehensive operations
                var contextMenu = new ContextMenuStrip();
                
                // Track this as the current active menu
                currentActiveMenu = contextMenu;
                
                // Determine what types of items we have
                bool hasFiles = selectedPaths.Any(path => File.Exists(path));
                bool hasFolders = selectedPaths.Any(path => Directory.Exists(path));
                bool allFiles = selectedPaths.All(path => File.Exists(path));
                bool allFolders = selectedPaths.All(path => Directory.Exists(path));
                
                // Open (only for single item type selections)
                if (selectedPaths.Count == 1 || allFiles || allFolders)
                {
                    var openItem = new ToolStripMenuItem("Open");
                    openItem.Click += (s, e) => {
                        foreach (string path in selectedPaths)
                        {
                            try
                            {
                                Process.Start(new ProcessStartInfo
                                {
                                    FileName = path,
                                    UseShellExecute = true
                                });
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show($"Error opening '{Path.GetFileName(path)}': {ex.Message}", 
                                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                break; // Stop on first error to avoid spam
                            }
                        }
                    };
                    contextMenu.Items.Add(openItem);
                    contextMenu.Items.Add(new ToolStripSeparator());
                }
                
                // Edit (only for files)
                if (hasFiles && selectedPaths.Count <= 5) // Limit to reasonable number
                {
                    var editItem = new ToolStripMenuItem("Edit");
                    editItem.Click += (s, e) => {
                        foreach (string path in selectedPaths.Where(p => File.Exists(p)))
                        {
                            try
                            {
                                Process.Start(new ProcessStartInfo
                                {
                                    FileName = "notepad.exe",
                                    Arguments = $"\"{path}\"",
                                    UseShellExecute = true
                                });
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show($"Error editing '{Path.GetFileName(path)}': {ex.Message}", 
                                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                break;
                            }
                        }
                    };
                    contextMenu.Items.Add(editItem);
                }
                
                // Send To submenu
                var sendToItem = new ToolStripMenuItem("Send To");
                var desktopItem = new ToolStripMenuItem("Desktop (create shortcut)");
                desktopItem.Click += (s, e) => {
                    string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                    MessageBox.Show($"Would create shortcuts for {selectedPaths.Count} items on Desktop", 
                        "Send To Desktop", MessageBoxButtons.OK, MessageBoxIcon.Information);
                };
                
                var documentsItem = new ToolStripMenuItem("My Documents");
                documentsItem.Click += (s, e) => {
                    string docsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                    MessageBox.Show($"Would copy {selectedPaths.Count} items to My Documents", 
                        "Send To Documents", MessageBoxButtons.OK, MessageBoxIcon.Information);
                };
                
                sendToItem.DropDownItems.AddRange(new[] { desktopItem, documentsItem });
                contextMenu.Items.Add(sendToItem);
                contextMenu.Items.Add(new ToolStripSeparator());
                
                // Cut
                var cutItem = new ToolStripMenuItem("Cut");
                cutItem.Click += (s, e) => {
                    CutFilesToClipboard(selectedPaths);
                };
                
                // Copy
                var copyItem = new ToolStripMenuItem("Copy");
                copyItem.Click += (s, e) => {
                    CopyFilesToClipboard(selectedPaths);
                };
                
                contextMenu.Items.AddRange(new ToolStripItem[] {
                    cutItem,
                    copyItem,
                    new ToolStripSeparator()
                });
                
                // Create shortcut (for files and folders)
                var shortcutItem = new ToolStripMenuItem("Create Shortcut");
                shortcutItem.Click += (s, e) => {
                    MessageBox.Show($"Would create shortcuts for {selectedPaths.Count} items in the current folder", 
                        "Create Shortcut", MessageBoxButtons.OK, MessageBoxIcon.Information);
                };
                contextMenu.Items.Add(shortcutItem);
                contextMenu.Items.Add(new ToolStripSeparator());
                
                // Delete
                var deleteItem = new ToolStripMenuItem("Delete");
                deleteItem.Click += (s, e) => {
                    deleteCallback?.Invoke();
                };
                contextMenu.Items.Add(deleteItem);
                
                // Rename (only for single item)
                if (selectedPaths.Count == 1)
                {
                    var renameItem = new ToolStripMenuItem("Rename");
                    renameItem.Click += (s, e) => {
                        renameCallback?.Invoke();
                    };
                    contextMenu.Items.Add(renameItem);
                }
                
                contextMenu.Items.Add(new ToolStripSeparator());
                
                // Properties
                var propertiesItem = new ToolStripMenuItem("Properties");
                propertiesItem.Click += (s, e) => {
                    ShowMultipleFileProperties(selectedPaths);
                };
                contextMenu.Items.Add(propertiesItem);
                
                // Handle menu closed event to clear tracking
                contextMenu.Closed += (s, e) => {
                    if (currentActiveMenu == contextMenu)
                        currentActiveMenu = null;
                };
                
                Debug.WriteLine($"Showing enhanced multiple items context menu with {contextMenu.Items.Count} items for {selectedPaths.Count} files");
                contextMenu.Show(screenPoint);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error showing multiple items context menu: {ex.Message}");
                // Fallback to simple message
                MessageBox.Show($"Right-clicked on {selectedPaths.Count} items\n\nAvailable operations:\n- Copy\n- Cut\n- Delete\n- Properties", 
                    "Multiple Items Selected", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void ShowNativeMultipleItemsContextMenu(List<string> selectedPaths, int x, int y)
        {
            Debug.WriteLine($"ShowNativeMultipleItemsContextMenu called for {selectedPaths.Count} items at ({x}, {y})");
            
            // Cancel any existing context menu first
            CancelCurrentContextMenu();
            
            try
            {
                // Try the COM interface approach for multiple items
                if (TryShowMultipleItemsComContextMenu(selectedPaths, x, y))
                {
                    Debug.WriteLine("Multiple items COM context menu succeeded");
                    return;
                }

                Debug.WriteLine("Multiple items COM context menu failed, trying simple fallback");
                // Fallback: Use the original custom context menu
                ShowMultipleItemsContextMenu(null!, selectedPaths, new Point(x, y));
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Exception in ShowNativeMultipleItemsContextMenu: {ex.Message}");
                // Fallback to custom menu
                ShowMultipleItemsContextMenu(null!, selectedPaths, new Point(x, y));
            }
        }

        private bool TryShowMultipleItemsComContextMenu(List<string> filePaths, int x, int y)
        {
            Debug.WriteLine($"TryShowMultipleItemsComContextMenu for {filePaths.Count} items");
            
            if (filePaths.Count == 0) return false;
            
            // For now, disable multiple items COM context menu due to stability issues
            // Fall back to custom menu which is safer and more reliable
            Debug.WriteLine("Multiple items COM context menu temporarily disabled for stability - using custom menu");
            return false;
        }
    }
}