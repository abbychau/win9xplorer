using System.ComponentModel;
using System.Diagnostics;

namespace win9xplorer
{
    public partial class Form1 : Form
    {
        private readonly NavigationManager navigationManager;
        private readonly FileSystemManager fileSystemManager;
        private readonly IconManager iconManager;
        private readonly ContextMenuManager contextMenuManager;
        private readonly BookmarkManager bookmarkManager;
        private readonly AddressBarSuggestionManager suggestionManager;
        private readonly RegistrySettingsManager registryManager;
        private System.Windows.Forms.Timer? resizeTimer;
        private System.Windows.Forms.Timer? splitterTimer;
        private System.Windows.Forms.Timer? toolStripTimer;
        
        public Form1()
        {
            InitializeComponent();
            
            // Initialize managers
            iconManager = new IconManager(imageListSmall, imageListLarge);
            fileSystemManager = new FileSystemManager(iconManager);
            navigationManager = new NavigationManager();
            contextMenuManager = new ContextMenuManager();
            bookmarkManager = new BookmarkManager();
            suggestionManager = new AddressBarSuggestionManager();
            registryManager = new RegistrySettingsManager();
            
            // Set up bookmark manager tree refresh callback
            bookmarkManager.SetTreeViewRefreshCallback(RefreshFavoritesTreeView);
            
            SetupImageLists();
            SetupContextMenu();
            SetupMenuBar();
            SetClassicWindowsStyle();
            SetupToolStripEvents();
            
            // Load settings from registry
            LoadAllSettings();
        }
        
        private void SetupToolStripEvents()
        {
            // Add events to save toolbar positions when moved
            toolStrip.LocationChanged += (s, e) => SaveToolStripPositions();
            addressStrip.LocationChanged += (s, e) => SaveToolStripPositions();
            
            // Add splitter moved event
            splitContainer.SplitterMoved += (s, e) => 
            {
                // Use a timer to avoid saving on every splitter move event during dragging
                splitterTimer?.Stop();
                splitterTimer = new System.Windows.Forms.Timer();
                splitterTimer.Interval = 300; // Save 300ms after splitter stops
                splitterTimer.Tick += (sender, args) =>
                {
                    registryManager.SaveWindowSettings(this, splitContainer, listView);
                    splitterTimer?.Stop();
                    splitterTimer?.Dispose();
                    splitterTimer = null;
                };
                splitterTimer.Start();
            };
        }
        
        private void SaveToolStripPositions()
        {
            // Use a timer to avoid excessive saves during toolbar dragging
            toolStripTimer?.Stop();
            toolStripTimer = new System.Windows.Forms.Timer();
            toolStripTimer.Interval = 500; // Save 500ms after movement stops
            toolStripTimer.Tick += (s, args) =>
            {
                registryManager.SaveToolStripSettings(toolStrip, addressStrip, txtAddress);
                toolStripTimer?.Stop();
                toolStripTimer?.Dispose();
                toolStripTimer = null;
            };
            toolStripTimer.Start();
        }
        
        private void LoadAllSettings()
        {
            try
            {
                Debug.WriteLine("Loading application settings from registry...");
                
                // Load and apply window settings
                var windowSettings = registryManager.LoadWindowSettings();
                registryManager.ApplyWindowSettings(this, splitContainer, listView, windowSettings);
                
                // Load and apply toolbar settings
                var toolbarSettings = registryManager.LoadToolStripSettings();
                registryManager.ApplyToolStripSettings(toolStrip, addressStrip, txtAddress, toolbarSettings);
                
                // Load and apply font settings
                var fontSettings = registryManager.LoadFontSettings();
                registryManager.ApplyFontSettings(treeView, listView, fontSettings);
                
                // Update view button states based on loaded settings
                UpdateViewButtonStates();
                
                // Update tree view visibility button state
                btnToggleTreeView.Checked = !splitContainer.Panel1Collapsed;
                
                // Navigate to the last used path
                if (!string.IsNullOrEmpty(windowSettings.LastPath))
                {
                    // Use BeginInvoke to ensure this happens after the form is fully loaded
                    this.BeginInvoke(new Action(() => {
                        if (windowSettings.LastPath == "My Computer")
                        {
                            NavigateToMyComputer();
                        }
                        else if (windowSettings.LastPath == "Favorites")
                        {
                            // Navigate to Favorites
                            foreach (TreeNode node in treeView.Nodes)
                            {
                                if (node.Tag?.ToString() == "favorites")
                                {
                                    treeView.SelectedNode = node;
                                    break;
                                }
                            }
                        }
                        else if (Directory.Exists(windowSettings.LastPath))
                        {
                            NavigateToPath(windowSettings.LastPath);
                        }
                        else
                        {
                            // Fallback to My Computer if last path doesn't exist
                            NavigateToMyComputer();
                        }
                    }));
                }
                
                Debug.WriteLine("Application settings loaded successfully");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error loading application settings: {ex.Message}");
            }
        }
        
        private void SaveAllSettings()
        {
            try
            {
                Debug.WriteLine("Saving application settings to registry...");
                
                // Save window settings
                registryManager.SaveWindowSettings(this, splitContainer, listView);
                
                // Save toolbar settings
                registryManager.SaveToolStripSettings(toolStrip, addressStrip, txtAddress);
                
                // Save font settings
                registryManager.SaveFontSettings(treeView.Font, listView.Font);
                
                // Bookmarks are saved automatically by BookmarkManager
                
                Debug.WriteLine("Application settings saved successfully");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error saving application settings: {ex.Message}");
            }
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            // Handle Alt+D to focus address bar
            if (keyData == (Keys.Alt | Keys.D))
            {
                FocusAddressBar();
                return true;
            }
            
            // Handle Alt+S to focus tree view (Sidebar/Folders)
            if (keyData == (Keys.Alt | Keys.S))
            {
                FocusTreeView();
                return true;
            }
            
            // Handle Ctrl+D to focus main list view (Directory)
            if (keyData == (Keys.Control | Keys.D))
            {
                FocusListView();
                return true;
            }
            
            // Handle F2 to rename selected item in ListView
            if (keyData == Keys.F2)
            {
                if (listView.Focused && listView.SelectedItems.Count == 1)
                {
                    StartRenameSelectedItem();
                    return true;
                }
            }
            
            // Handle Delete key to delete selected items in ListView
            if (keyData == Keys.Delete)
            {
                if (listView.Focused && listView.SelectedItems.Count > 0)
                {
                    DeleteSelectedItems();
                    return true;
                }
            }
            
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            // Handle Alt+D to focus address bar (backup method)
            if (e.Alt && e.KeyCode == Keys.D)
            {
                FocusAddressBar();
                e.Handled = true;
            }
            // Handle Alt+S to focus tree view (backup method)
            else if (e.Alt && e.KeyCode == Keys.S)
            {
                FocusTreeView();
                e.Handled = true;
            }
            // Handle Ctrl+D to focus list view (backup method)
            else if (e.Control && e.KeyCode == Keys.D)
            {
                FocusListView();
                e.Handled = true;
            }
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            // Make address bar elastic by adjusting its size when form resizes
            if (addressStrip != null && txtAddress != null)
            {
                try
                {
                    // Calculate available space for address textbox more carefully
                    int totalStripWidth = addressStrip.Width;
                    int usedWidth = lblAddress.Width + btnBookmark.Width;
                    
                    // Add some padding and ensure we don't make it too small
                    int paddingAndMargin = 60; // Increased padding for safety
                    int availableWidth = totalStripWidth - usedWidth - paddingAndMargin;
                    
                    // Set minimum and maximum width constraints
                    int minWidth = 150; // Minimum useful width
                    int maxWidth = 500; // Maximum reasonable width
                    
                    // Ensure the width stays within reasonable bounds
                    availableWidth = Math.Max(minWidth, Math.Min(availableWidth, maxWidth));
                    
                    // Only resize if we have a reasonable amount of space
                    if (availableWidth >= minWidth && totalStripWidth > 0)
                    {
                        txtAddress.Size = new Size(availableWidth, txtAddress.Size.Height);
                    }
                    
                    Debug.WriteLine($"Address bar resized to: {availableWidth}px (Strip: {totalStripWidth}px, Used: {usedWidth}px)");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Error resizing address bar: {ex.Message}");
                }
            }
            
            // Save window settings on resize (but only if window is not minimized to avoid excessive saves)
            if (this.WindowState != FormWindowState.Minimized)
            {
                // Use a timer to avoid saving on every resize event during dragging
                resizeTimer?.Stop();
                resizeTimer = new System.Windows.Forms.Timer();
                resizeTimer.Interval = 500; // Save 500ms after resize stops
                resizeTimer.Tick += (s, args) =>
                {
                    registryManager.SaveWindowSettings(this, splitContainer, listView);
                    resizeTimer?.Stop();
                    resizeTimer?.Dispose();
                    resizeTimer = null;
                };
                resizeTimer.Start();
            }
        }

        private void FocusAddressBar()
        {
            txtAddress.Focus();
            txtAddress.SelectAll();
        }

        private void FocusTreeView()
        {
            if (!splitContainer.Panel1Collapsed)
            {
                treeView.Focus();
                // Ensure a node is selected for keyboard navigation
                if (treeView.SelectedNode == null && treeView.Nodes.Count > 0)
                {
                    treeView.SelectedNode = treeView.Nodes[0];
                }
            }
        }

        private void FocusListView()
        {
            listView.Focus();
            // If there are items and none are selected, select the first one
            if (listView.Items.Count > 0 && listView.SelectedItems.Count == 0)
            {
                listView.Items[0].Selected = true;
                listView.Items[0].Focused = true;
            }
        }

        private void SetClassicWindowsStyle()
        {
            // Set classic Windows 95/98 colors and appearance
            this.BackColor = SystemColors.Control;
            treeView.BackColor = Color.White;
            listView.BackColor = Color.White;
            
            // Set the form icon
            this.Icon = iconManager.CreateApplicationIcon();
            
            // Make sure the split container has the classic look
            splitContainer.BackColor = SystemColors.Control;
            splitContainer.Panel1.BackColor = SystemColors.Control;
            splitContainer.Panel2.BackColor = SystemColors.Control;
            
            // Set up coolbar styling for Windows 9x look
            toolStrip.GripStyle = ToolStripGripStyle.Visible;
            addressStrip.GripStyle = ToolStripGripStyle.Visible;
        }

        private void SetupMenuBar()
        {
            // Setup File menu
            var newMenuItem = new ToolStripMenuItem("&New");
            var newFolderItem = new ToolStripMenuItem("&Folder", null, (s, e) => CreateNewFolder());
            newMenuItem.DropDownItems.Add(newFolderItem);
            
            var closeMenuItem = new ToolStripMenuItem("&Close", null, (s, e) => this.Close());
            var exitMenuItem = new ToolStripMenuItem("E&xit", null, (s, e) => Application.Exit());

            fileMenu.DropDownItems.AddRange(new ToolStripItem[] {
                newMenuItem,
                new ToolStripSeparator(),
                closeMenuItem,
                new ToolStripSeparator(),
                exitMenuItem
            });

            // Setup Edit menu
            var selectAllMenuItem = new ToolStripMenuItem("Select &All", null, (s, e) => SelectAllItems());
            var copyMenuItem = new ToolStripMenuItem("&Copy", null, (s, e) => CopySelectedItems());
            var cutMenuItem = new ToolStripMenuItem("Cu&t", null, (s, e) => CutSelectedItems());
            var pasteMenuItem = new ToolStripMenuItem("&Paste", null, (s, e) => PasteItems());
            var deleteMenuItem = new ToolStripMenuItem("&Delete", null, (s, e) => DeleteSelectedItems());
            var renameMenuItem = new ToolStripMenuItem("&Rename", null, (s, e) => StartRenameSelectedItem());

            editMenu.DropDownItems.AddRange(new ToolStripItem[] {
                selectAllMenuItem,
                new ToolStripSeparator(),
                copyMenuItem,
                cutMenuItem,
                pasteMenuItem,
                new ToolStripSeparator(),
                deleteMenuItem,
                renameMenuItem
            });

            // Setup View menu
            var toolbarMenuItem = new ToolStripMenuItem("&Toolbar", null, (s, e) => ToggleToolbar()) { Checked = true };
            var statusBarMenuItem = new ToolStripMenuItem("&Status Bar", null, (s, e) => ToggleStatusBar()) { Checked = true };
            var explorerBarMenuItem = new ToolStripMenuItem("&Explorer Bar");
            var foldersMenuItem = new ToolStripMenuItem("&Folders", null, (s, e) => ToggleTreeView()) { Checked = true };
            explorerBarMenuItem.DropDownItems.Add(foldersMenuItem);
            
            var largeIconsMenuItem = new ToolStripMenuItem("Lar&ge Icons", null, (s, e) => SetViewMode(View.LargeIcon));
            var smallIconsMenuItem = new ToolStripMenuItem("S&mall Icons", null, (s, e) => SetViewMode(View.SmallIcon));
            var listMenuItem = new ToolStripMenuItem("&List", null, (s, e) => SetViewMode(View.List));
            var detailsMenuItem = new ToolStripMenuItem("&Details", null, (s, e) => SetViewMode(View.Details));

            viewMenu.DropDownItems.AddRange(new ToolStripItem[] {
                toolbarMenuItem,
                statusBarMenuItem,
                explorerBarMenuItem,
                new ToolStripSeparator(),
                largeIconsMenuItem,
                smallIconsMenuItem,
                listMenuItem,
                detailsMenuItem,
                new ToolStripSeparator(),
                new ToolStripMenuItem("&Refresh", null, (s, e) => BtnRefresh_Click(s!, e))
            });

            // Setup Go menu
            var backMenuItem = new ToolStripMenuItem("&Back", null, (s, e) => BtnBack_Click(s!, e));
            var forwardMenuItem = new ToolStripMenuItem("&Forward", null, (s, e) => BtnForward_Click(s!, e));
            var upMenuItem = new ToolStripMenuItem("&Up One Level", null, (s, e) => BtnUp_Click(s!, e));
            
            goMenu.DropDownItems.AddRange(new ToolStripItem[] {
                backMenuItem,
                forwardMenuItem,
                upMenuItem,
                new ToolStripSeparator(),
                new ToolStripMenuItem("My &Computer", null, (s, e) => NavigateToMyComputer()),
                new ToolStripMenuItem("My &Documents", null, (s, e) => NavigateToPath(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)))
            });

            // Setup Favorites menu
            var addToFavoritesMenuItem = new ToolStripMenuItem("&Add to Favorites", null, (s, e) => AddCurrentPathToBookmarks());
            favoritesMenu.DropDownItems.Add(addToFavoritesMenuItem);
            
            // Populate with existing bookmarks
            RefreshFavoritesMenu();

            // Setup Tools menu
            var optionsMenuItem = new ToolStripMenuItem("&Options...", null, (s, e) => ShowOptions());
            var exportSettingsMenuItem = new ToolStripMenuItem("&Export Settings...", null, (s, e) => ExportSettings());
            var importSettingsMenuItem = new ToolStripMenuItem("&Import Settings...", null, (s, e) => ImportSettings());
            var resetSettingsMenuItem = new ToolStripMenuItem("&Reset All Settings...", null, (s, e) => ResetAllSettings());
            toolsMenu.DropDownItems.AddRange(new ToolStripItem[] {
                optionsMenuItem,
                new ToolStripSeparator(),
                exportSettingsMenuItem,
                importSettingsMenuItem,
                new ToolStripSeparator(),
                resetSettingsMenuItem
            });

            // Setup Help menu
            var aboutMenuItem = new ToolStripMenuItem("&About Windows Explorer", null, (s, e) => ShowAbout());
            var testIconsMenuItem = new ToolStripMenuItem("&Test Windows XP Icons", null, (s, e) => iconManager.TestWindowsXPIconLoading());
            var refreshIconsMenuItem = new ToolStripMenuItem("&Refresh Toolbar Icons", null, (s, e) => RefreshToolbarIcons());
            
            helpMenu.DropDownItems.AddRange(new ToolStripItem[] {
                testIconsMenuItem,
                refreshIconsMenuItem,
                new ToolStripSeparator(),
                aboutMenuItem
            });
        }

        private void SetupContextMenu()
        {
            contextMenuManager.SetupContextMenu(listView);
            contextMenuManager.SetupTreeViewContextMenu(treeView);
            contextMenuManager.SetupRenameCallback(StartRenameSelectedItem);
            contextMenuManager.SetupDeleteCallback(DeleteSelectedItems);
            
            // Add mouse event handlers - MouseDown for cancellation, MouseUp for showing menu
            listView.MouseDown += ListView_MouseDown;
            listView.MouseUp += ListView_MouseUp;
        }

        private void ListView_MouseDown(object sender, MouseEventArgs e)
        {
            Debug.WriteLine($"ListView_MouseDown called - Button: {e.Button}, Location: {e.Location}");
            
            // Let the context menu manager handle right-click down (for cancellation)
            if (e.Button == MouseButtons.Right)
            {
                contextMenuManager.OnListViewMouseDown(e);
            }
        }

        // Menu event handlers
        private void CreateNewFolder()
        {
            // Implementation for creating new folder
            MessageBox.Show("New Folder functionality would be implemented here.", "New Folder", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void SelectAllItems()
        {
            foreach (ListViewItem item in listView.Items)
            {
                item.Selected = true;
            }
        }

        private void CopySelectedItems()
        {
            try
            {
                var selectedPaths = new List<string>();
                foreach (ListViewItem item in listView.SelectedItems)
                {
                    string itemPath = item.Tag?.ToString() ?? "";
                    if (!string.IsNullOrEmpty(itemPath))
                    {
                        selectedPaths.Add(itemPath);
                    }
                }
                
                if (selectedPaths.Count == 0)
                {
                    MessageBox.Show("No items selected to copy.", "Copy", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                
                // Create a StringCollection with file paths
                var files = new System.Collections.Specialized.StringCollection();
                files.AddRange(selectedPaths.ToArray());
                
                // Copy to clipboard
                Clipboard.SetFileDropList(files);
                
                MessageBox.Show($"Copied {selectedPaths.Count} item(s) to clipboard.", "Copy", 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error copying items: {ex.Message}", "Copy Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void CutSelectedItems()
        {
            try
            {
                var selectedPaths = new List<string>();
                foreach (ListViewItem item in listView.SelectedItems)
                {
                    string itemPath = item.Tag?.ToString() ?? "";
                    if (!string.IsNullOrEmpty(itemPath))
                    {
                        selectedPaths.Add(itemPath);
                    }
                }
                
                if (selectedPaths.Count == 0)
                {
                    MessageBox.Show("No items selected to cut.", "Cut", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                
                // Create a StringCollection with file paths
                var files = new System.Collections.Specialized.StringCollection();
                files.AddRange(selectedPaths.ToArray());
                
                // Copy to clipboard with cut format
                var dataObject = new DataObject();
                dataObject.SetFileDropList(files);
                
                // Add a special format to indicate cut (not copy)
                byte[] moveEffect = BitConverter.GetBytes(2); // DROPEFFECT_MOVE
                dataObject.SetData("Preferred DropEffect", moveEffect);
                
                Clipboard.SetDataObject(dataObject, true);
                
                MessageBox.Show($"Cut {selectedPaths.Count} item(s) to clipboard.", "Cut", 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error cutting items: {ex.Message}", "Cut Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void PasteItems()
        {
            try
            {
                if (!Clipboard.ContainsFileDropList())
                {
                    MessageBox.Show("No files in clipboard to paste.", "Paste", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                
                var files = Clipboard.GetFileDropList();
                if (files.Count == 0)
                {
                    MessageBox.Show("No files in clipboard to paste.", "Paste", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                
                string targetDirectory = navigationManager.CurrentPath;
                if (string.IsNullOrEmpty(targetDirectory))
                {
                    MessageBox.Show("Cannot paste files here.", "Paste", 
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
                // Check if this is a cut operation
                bool isCutOperation = false;
                try
                {
                    var dataObject = Clipboard.GetDataObject();
                    if (dataObject?.GetDataPresent("Preferred DropEffect") == true)
                    {
                        var dropEffect = (byte[])dataObject.GetData("Preferred DropEffect");
                        isCutOperation = BitConverter.ToInt32(dropEffect, 0) == 2; // DROPEFFECT_MOVE
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Error checking drop effect: {ex.Message}");
                }
                
                string operation = isCutOperation ? "move" : "copy";
                var result = MessageBox.Show($"Are you sure you want to {operation} {files.Count} item(s) to this location?", 
                    $"{(isCutOperation ? "Move" : "Copy")} Files", 
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                
                if (result == DialogResult.Yes)
                {
                    int successCount = 0;
                    foreach (string file in files)
                    {
                        try
                        {
                            string fileName = Path.GetFileName(file);
                            string destPath = Path.Combine(targetDirectory, fileName);
                            
                            if (Directory.Exists(file))
                            {
                                // Handle directory
                                if (isCutOperation)
                                {
                                    Directory.Move(file, destPath);
                                }
                                else
                                {
                                    CopyDirectory(file, destPath);
                                }
                            }
                            else if (File.Exists(file))
                            {
                                // Handle file
                                if (isCutOperation)
                                {
                                    File.Move(file, destPath);
                                }
                                else
                                {
                                    File.Copy(file, destPath, true);
                                }
                            }
                            
                            successCount++;
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Error {operation}ing '{Path.GetFileName(file)}': {ex.Message}", 
                                $"{(isCutOperation ? "Move" : "Copy")} Error", 
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                    
                    if (successCount > 0)
                    {
                        // Refresh the current view
                        BtnRefresh_Click(this, EventArgs.Empty);
                        
                        MessageBox.Show($"Successfully {operation}ed {successCount} item(s).", 
                            $"{(isCutOperation ? "Move" : "Copy")} Complete", 
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error pasting items: {ex.Message}", "Paste Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void CopyDirectory(string sourceDir, string destDir)
        {
            // Create destination directory
            Directory.CreateDirectory(destDir);
            
            // Copy files
            foreach (string filePath in Directory.GetFiles(sourceDir))
            {
                string fileName = Path.GetFileName(filePath);
                string destFile = Path.Combine(destDir, fileName);
                File.Copy(filePath, destFile, true);
            }
            
            // Copy subdirectories recursively
            foreach (string dirPath in Directory.GetDirectories(sourceDir))
            {  
                string dirName = Path.GetFileName(dirPath);
                string destSubDir = Path.Combine(destDir, dirName);
                CopyDirectory(dirPath, destSubDir);
            }
        }

        private void ToggleToolbar()
        {
            toolStrip.Visible = !toolStrip.Visible;
            addressStrip.Visible = !addressStrip.Visible;
            
            // Update the menu item check state
            if (viewMenu.DropDownItems.Count > 0 && viewMenu.DropDownItems[0] is ToolStripMenuItem toolbarMenuItem)
            {
                toolbarMenuItem.Checked = toolStrip.Visible;
            }
            
            // Save toolbar visibility settings
            SaveAllSettings();
        }

        private void ToggleStatusBar()
        {
            statusStrip.Visible = !statusStrip.Visible;
            
            // Update the menu item check state
            if (viewMenu.DropDownItems.Count > 1 && viewMenu.DropDownItems[1] is ToolStripMenuItem statusBarMenuItem)
            {
                statusBarMenuItem.Checked = statusStrip.Visible;
            }
            
            // Save status bar visibility settings
            SaveAllSettings();
        }

        private void ToggleTreeView()
        {
            splitContainer.Panel1Collapsed = !splitContainer.Panel1Collapsed;
            btnToggleTreeView.Checked = !splitContainer.Panel1Collapsed;
            
            // Update the View menu item
            if (viewMenu.DropDownItems.Count > 2)
            {
                var explorerBarItem = viewMenu.DropDownItems[2] as ToolStripMenuItem;
                if (explorerBarItem?.DropDownItems.Count > 0)
                {
                    var foldersItem = explorerBarItem.DropDownItems[0] as ToolStripMenuItem;
                    if (foldersItem != null)
                    {
                        foldersItem.Checked = !splitContainer.Panel1Collapsed;
                    }
                }
            }
            
            // Save tree view visibility settings
            registryManager.SaveWindowSettings(this, splitContainer, listView);
        }

        private void BtnToggleTreeView_Click(object sender, EventArgs e)
        {
            ToggleTreeView();
        }

        private void NavigateToMyComputer()
        {
            // Navigate to My Computer root and select it in the tree view
            foreach (TreeNode node in treeView.Nodes)
            {
                if (node.Tag?.ToString() == "computer")
                {
                    treeView.SelectedNode = node;
                    // The TreeView_AfterSelect event will handle loading the contents
                    break;
                }
            }
        }

        private void ShowOptions()
        {
            // Show options dialog
            using (var optionsDialog = new OptionsDialog(treeView.Font, listView.Font))
            {
                if (optionsDialog.ShowDialog(this) == DialogResult.OK)
                {
                    // Apply the selected fonts
                    treeView.Font = optionsDialog.TreeViewFont;
                    listView.Font = optionsDialog.ListViewFont;
                    
                    // Save font settings to registry
                    registryManager.SaveFontSettings(treeView.Font, listView.Font);
                    
                    Debug.WriteLine("Font settings applied and saved to registry");
                }
            }
        }

        private void ShowAbout()
        {
            MessageBox.Show($"Windows 9x Explorer Clone\n\nBuilt with .NET 8\nVersion 1.0", "About Windows Explorer", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Debug.WriteLine("=== Form1_Load Started ===");
            
            LoadDrives();
            
            // Navigate to My Computer instead of Desktop by default
            NavigateToMyComputer();
            
            // Set up toolbar icons after form loads
            SetupToolbarIcons();
            
            // Initialize view button states
            UpdateViewButtonStates();
            
            // Setup address bar suggestions
            suggestionManager.SetupSuggestions(this, txtAddress);
            
            Debug.WriteLine("=== Form1_Load Completed ===");
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            // Save all settings before closing
            SaveAllSettings();
            
            // Clean up suggestion manager
            suggestionManager.Dispose();
        }

        private void SetupToolbarIcons()
        {
            iconManager.SetupToolbarIcons(btnBack, btnForward, btnUp, btnRefresh, 
                btnViewLargeIcons, btnViewSmallIcons, btnViewList, btnViewDetails);
            
            // Setup TreeView toggle button icon
            SetupTreeViewToggleIcon();
            
            // Setup bookmark button icon
            SetupBookmarkIcon();
            
            // Setup Navigate to Folder button icon
            SetupNavigateToFolderIcon();
            
            // Update view button states based on current view
            UpdateViewButtonStates();
        }

        private void SetupTreeViewToggleIcon()
        {
            // Create a simple folder tree icon for the toggle button
            var folderIcon = iconManager.CreateViewIcon("folders", 16);
            if (folderIcon != null)
            {
                btnToggleTreeView.Image = folderIcon;
            }
        }

        private void SetupBookmarkIcon()
        {
            // Create a star icon for the bookmark button
            var starIcon = iconManager.CreateViewIcon("star", 16);
            if (starIcon != null)
            {
                btnBookmark.Image = starIcon;
            }
            
            // Update the bookmark button state
            UpdateBookmarkButtonState();
        }

        private void SetupNavigateToFolderIcon()
        {
            // Create a navigate icon for the navigate to folder button
            var navigateIcon = iconManager.CreateViewIcon("navigate", 16);
            if (navigateIcon != null)
            {
                btnNavigateToFolder.Image = navigateIcon;
            }
        }

        private void UpdateBookmarkButtonState()
        {
            string currentPath = string.IsNullOrEmpty(navigationManager.CurrentPath) ? "My Computer" : navigationManager.CurrentPath;
            bool isBookmarked = bookmarkManager.IsBookmarked(currentPath);
            
            btnBookmark.Checked = isBookmarked;
            btnBookmark.ToolTipText = isBookmarked ? "Remove from Favorites" : "Add to Favorites";
        }

        private void BtnBookmark_Click(object sender, EventArgs e)
        {
            string currentPath = string.IsNullOrEmpty(navigationManager.CurrentPath) ? "My Computer" : navigationManager.CurrentPath;
            
            if (bookmarkManager.IsBookmarked(currentPath))
            {
                bookmarkManager.RemoveBookmark(currentPath);
            }
            else
            {
                bookmarkManager.AddBookmark(currentPath);
            }
            
            UpdateBookmarkButtonState();
            RefreshFavoritesMenu();
        }

        private void AddCurrentPathToBookmarks()
        {
            string currentPath = string.IsNullOrEmpty(navigationManager.CurrentPath) ? "My Computer" : navigationManager.CurrentPath;
            bookmarkManager.AddBookmark(currentPath);
            UpdateBookmarkButtonState();
            RefreshFavoritesMenu();
        }

        private void RefreshFavoritesMenu()
        {
            bookmarkManager.PopulateFavoritesMenu(favoritesMenu, NavigateToBookmarkedPath);
        }

        private void NavigateToBookmarkedPath(string path)
        {
            if (path == "My Computer")
            {
                NavigateToMyComputer();
            }
            else
            {
                NavigateToPath(path, false); // Don't sync tree view for bookmark navigation from menu
            }
        }

        /// <summary>
        /// Public method to refresh toolbar icons - useful for debugging or runtime icon changes
        /// </summary>
        public void RefreshToolbarIcons()
        {
            Debug.WriteLine("=== Manual Toolbar Icons Refresh Requested ===");
            SetupToolbarIcons();
            this.Refresh(); // Refresh the form to show updated icons
        }

        // Event handlers for View As buttons
        private void BtnViewLargeIcons_Click(object sender, EventArgs e)
        {
            SetViewMode(View.LargeIcon);
        }

        private void BtnViewSmallIcons_Click(object sender, EventArgs e)
        {
            SetViewMode(View.SmallIcon);
        }

        private void BtnViewList_Click(object sender, EventArgs e)
        {
            SetViewMode(View.List);
        }

        private void BtnViewDetails_Click(object sender, EventArgs e)
        {
            SetViewMode(View.Details);
        }

        private void SetViewMode(View viewMode)
        {
            listView.View = viewMode;
            
            if (viewMode == View.Details)
            {
                // Check if we're currently viewing My Computer
                if (string.IsNullOrEmpty(navigationManager.CurrentPath) && txtAddress.Text == "My Computer")
                {
                    // Reload My Computer contents with the new Details view
                    LoadMyComputerContents();
                }
                else if (txtAddress.Text == "Favorites")
                {
                    // Reload Favorites with details view
                    LoadFavoritesContents();
                }
                else
                {
                    // Regular directory - use standard details view
                    fileSystemManager.SetupDetailsView(listView);
                    fileSystemManager.LoadDirectoryContents(navigationManager.CurrentPath, listView);
                }
            }
            else
            {
                // For non-Details views, we may need to refresh special contents
                if (string.IsNullOrEmpty(navigationManager.CurrentPath))
                {
                    if (txtAddress.Text == "My Computer")
                    {
                        LoadMyComputerContents();
                    }
                    else if (txtAddress.Text == "Favorites")
                    {
                        LoadFavoritesContents();
                    }
                }
            }
            
            // Update button states to show which view is active
            UpdateViewButtonStates();
            
            // Save window settings to preserve view mode
            registryManager.SaveWindowSettings(this, splitContainer, listView);
        }

        private void UpdateViewButtonStates()
        {
            // Update button appearance based on current view mode
            btnViewLargeIcons.Checked = (listView.View == View.LargeIcon);
            btnViewSmallIcons.Checked = (listView.View == View.SmallIcon);
            btnViewList.Checked = (listView.View == View.List);
            btnViewDetails.Checked = (listView.View == View.Details);
        }

        private void SetupImageLists()
        {
            iconManager.SetupImageLists();
        }

        private void LoadDrives()
        {
            navigationManager.LoadDrives(treeView, imageListSmall, imageListLarge, iconManager);
            
            // Populate favorites initially
            RefreshFavoritesTreeView();
        }

        private void TreeView_BeforeExpand(object sender, TreeViewCancelEventArgs e)
        {
            if (e.Node?.Tag?.ToString() == "computer" || e.Node?.Tag?.ToString() == "favorites")
                return;

            // Clear dummy nodes and load actual folders
            fileSystemManager.ExpandTreeNode(e.Node!);
        }

        private void TreeView_AfterSelect(object sender, TreeViewEventArgs e)
        {
            string nodeTag = e.Node?.Tag?.ToString() ?? "";
            
            if (nodeTag == "computer")
            {
                // When "My Computer" is selected, show all drives in the main list view
                LoadMyComputerContents();
                return;
            }
            else if (nodeTag == "favorites")
            {
                // When "Favorites" is selected, show bookmarks in the main list view
                LoadFavoritesContents();
                return;
            }
            else if (nodeTag.StartsWith("bookmark:"))
            {
                // Handle bookmark navigation from tree view - don't sync tree since user already clicked in tree
                string bookmarkPath = nodeTag.Substring("bookmark:".Length);
                if (bookmarkPath == "My Computer")
                {
                    NavigateToMyComputer();
                }
                else if (Directory.Exists(bookmarkPath))
                {
                    NavigateToPath(bookmarkPath, false); // Don't sync tree view since user clicked from tree
                }
                return;
            }

            if (!string.IsNullOrEmpty(nodeTag) && Directory.Exists(nodeTag))
            {
                NavigateToPath(nodeTag, true); // Normal tree navigation should sync
            }
        }

        private void LoadMyComputerContents()
        {
            // Hide address bar suggestions when loading My Computer
            suggestionManager.HideSuggestions();
            
            // Set flag to prevent suggestions during programmatic navigation
            suggestionManager.SetNavigatingProgrammatically(true);
            
            navigationManager.CurrentPath = ""; // Clear current path for My Computer
            txtAddress.Text = "My Computer";
            
            // Reset the flag after setting text
            suggestionManager.SetNavigatingProgrammatically(false);

            // If we're in Details view, set up the My Computer specific columns
            if (listView.View == View.Details)
            {
                fileSystemManager.SetupMyComputerDetailsView(listView);
            }

            fileSystemManager.LoadMyComputerContents(listView, imageListSmall, imageListLarge);
            
            // Update status bar
            statusLabel.Text = $"{listView.Items.Count} drive(s)";
            
            // Update navigation buttons - disable back/forward/up for My Computer
            UpdateNavigationButtons();
            
            // Update window title to match address
            UpdateWindowTitle();
            
            // Update bookmark button state
            UpdateBookmarkButtonState();
            
            // Save current path
            SaveCurrentPath();
        }

        private void LoadFavoritesContents()
        {
            // Hide address bar suggestions when loading Favorites
            suggestionManager.HideSuggestions();
            
            // Set flag to prevent suggestions during programmatic navigation
            suggestionManager.SetNavigatingProgrammatically(true);
            
            navigationManager.CurrentPath = ""; // Clear current path for Favorites
            txtAddress.Text = "Favorites";
            
            // Reset the flag after setting text
            suggestionManager.SetNavigatingProgrammatically(false);

            listView.Clear();
            listView.View = View.LargeIcon;
            
            // Add columns for details view
            if (listView.View == View.Details)
            {
                listView.Columns.Clear();
                listView.Columns.Add("Name", 200);
                listView.Columns.Add("Location", 300);
                listView.Columns.Add("Date Added", 120);
            }

            // Load bookmarks into list view
            var bookmarks = bookmarkManager.GetBookmarks();
            foreach (var bookmark in bookmarks)
            {
                var item = new ListViewItem(bookmark.Name)
                {
                    Tag = bookmark.Path,
                    ImageKey = bookmark.Path == "My Computer" ? "computer" : "folder"
                };
                
                if (listView.View == View.Details)
                {
                    item.SubItems.Add(bookmark.Path);
                    item.SubItems.Add(bookmark.DateAdded.ToString("yyyy-MM-dd"));
                }
                
                listView.Items.Add(item);
            }
            
            // Update status bar
            statusLabel.Text = $"{listView.Items.Count} favorite(s)";
            
            // Update navigation buttons
            UpdateNavigationButtons();
            
            // Update window title
            UpdateWindowTitle();
            
            // Update bookmark button state
            UpdateBookmarkButtonState();
            
            // Save current path
            SaveCurrentPath();
        }

        private void NavigateToPath(string path)
        {
            NavigateToPath(path, true);
        }

        private void NavigateToPath(string path, bool syncTreeView)
        {
            try
            {
                // Hide address bar suggestions when navigating programmatically
                suggestionManager.HideSuggestions();
                
                // Set flag to prevent suggestions during programmatic navigation
                suggestionManager.SetNavigatingProgrammatically(true);
                
                navigationManager.CurrentPath = path;
                txtAddress.Text = path;
                
                // Reset the flag after setting text
                suggestionManager.SetNavigatingProgrammatically(false);
                
                // If we're switching from My Computer to a regular path and we're in Details view,
                // we need to reset the columns to the standard layout
                if (listView.View == View.Details)
                {
                    fileSystemManager.SetupDetailsView(listView);
                }
                
                fileSystemManager.LoadDirectoryContents(path, listView);
                
                // Add to navigation history
                navigationManager.AddToHistory(path);

                UpdateNavigationButtons();
                UpdateStatusBar();
                
                // Sync tree view selection only if requested
                if (syncTreeView)
                {
                    navigationManager.SyncTreeViewSelection(treeView, path);
                }
                
                // Update window title to match address
                UpdateWindowTitle();
                
                // Update bookmark button state
                UpdateBookmarkButtonState();
                
                // Save current path immediately for persistence
                SaveCurrentPath();
            }
            catch (Exception ex)
            {
                // Make sure to reset the flag even if an error occurs
                suggestionManager.SetNavigatingProgrammatically(false);
                MessageBox.Show($"Error navigating to {path}: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        
        private void SaveCurrentPath()
        {
            try
            {
                // Save just the current path setting without affecting other window settings
                using (var key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(@"SOFTWARE\win9xplorer\Window"))
                {
                    string currentPath = string.IsNullOrEmpty(navigationManager.CurrentPath) ? txtAddress.Text : navigationManager.CurrentPath;
                    key.SetValue("LastPath", currentPath);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error saving current path: {ex.Message}");
            }
        }

        private void RefreshFavoritesTreeView()
        {
            var bookmarks = bookmarkManager.GetBookmarks();
            navigationManager.PopulateFavoritesNode(treeView, bookmarks, iconManager);
        }

        private void StartRenameSelectedItem()
        {
            if (listView.SelectedItems.Count == 1)
            {
                var selectedItem = listView.SelectedItems[0];
                
                // Don't allow renaming in special views (My Computer, Favorites)
                if (string.IsNullOrEmpty(navigationManager.CurrentPath))
                {
                    if (txtAddress.Text == "My Computer")
                    {
                        MessageBox.Show("Cannot rename drives.", "Rename", 
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                    else if (txtAddress.Text == "Favorites")
                    {
                        MessageBox.Show("Use the Favorites menu to manage bookmarks.", "Rename", 
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                }
                
                // Start label editing for the selected item
                selectedItem.BeginEdit();
            }
        }

        private void DeleteSelectedItems()
        {
            try
            {
                if (listView.SelectedItems.Count == 0)
                {
                    MessageBox.Show("No items selected to delete.", "Delete", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Don't allow deleting in special views (My Computer, Favorites)
                if (string.IsNullOrEmpty(navigationManager.CurrentPath))
                {
                    if (txtAddress.Text == "My Computer")
                    {
                        MessageBox.Show("Cannot delete drives.", "Delete", 
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                    else if (txtAddress.Text == "Favorites")
                    {
                        MessageBox.Show("Use the Favorites menu to manage bookmarks.", "Delete", 
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                }

                // Get selected file paths
                var selectedPaths = new List<string>();
                foreach (ListViewItem item in listView.SelectedItems)
                {
                    string itemPath = item.Tag?.ToString() ?? "";
                    if (!string.IsNullOrEmpty(itemPath))
                    {
                        selectedPaths.Add(itemPath);
                    }
                }

                if (selectedPaths.Count == 0)
                {
                    MessageBox.Show("No valid items selected to delete.", "Delete", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Show confirmation dialog
                string message = selectedPaths.Count == 1 
                    ? $"Are you sure you want to delete '{Path.GetFileName(selectedPaths[0])}'?"
                    : $"Are you sure you want to delete these {selectedPaths.Count} items?";
                
                var result = MessageBox.Show(message, "Delete Files", 
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                
                if (result == DialogResult.Yes)
                {
                    int successCount = 0;
                    var errors = new List<string>();
                    
                    foreach (string filePath in selectedPaths)
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
                    
                    // Refresh the current view to reflect deletions
                    BtnRefresh_Click(this, EventArgs.Empty);
                    
                    // Show results
                    if (errors.Count == 0)
                    {
                        Debug.WriteLine($"Successfully deleted {successCount} items");
                    }
                    else if (successCount > 0)
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
            catch (Exception ex)
            {
                Debug.WriteLine($"Error in DeleteSelectedItems: {ex.Message}");
                MessageBox.Show($"An error occurred during deletion: {ex.Message}", 
                    "Delete Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnNavigateToFolder_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(navigationManager.CurrentPath))
            {
                MessageBox.Show("No folder is currently selected to navigate to in the tree.", 
                    "Navigate to Folder", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Find and select the current folder in the tree view
            navigationManager.SyncTreeViewSelection(treeView, navigationManager.CurrentPath);
            
            // Expand the tree to show the current folder
            ExpandTreeToPath(navigationManager.CurrentPath);
        }

        private void ExpandTreeToPath(string targetPath)
        {
            // Find the computer node first
            foreach (TreeNode computerNode in treeView.Nodes)
            {
                if (computerNode.Tag?.ToString() == "computer")
                {
                    // Find the drive that contains this path
                    foreach (TreeNode driveNode in computerNode.Nodes)
                    {
                        string driveTag = driveNode.Tag?.ToString() ?? "";
                        if (!string.IsNullOrEmpty(driveTag) && targetPath.StartsWith(driveTag, StringComparison.OrdinalIgnoreCase))
                        {
                            // Expand this drive
                            if (!driveNode.IsExpanded)
                            {
                                driveNode.Expand();
                            }
                            
                            // Recursively expand to find the target folder
                            ExpandToTargetPath(driveNode, targetPath);
                            break;
                        }
                    }
                    break;
                }
            }
        }

        private void ExpandToTargetPath(TreeNode parentNode, string targetPath)
        {
            foreach (TreeNode childNode in parentNode.Nodes)
            {
                string childTag = childNode.Tag?.ToString() ?? "";
                if (childTag == targetPath)
                {
                    // Found the target - select it
                    treeView.SelectedNode = childNode;
                    childNode.EnsureVisible();
                    return;
                }
                else if (!string.IsNullOrEmpty(childTag) && targetPath.StartsWith(childTag, StringComparison.OrdinalIgnoreCase))
                {
                    // This node is on the path to our target - expand it
                    if (!childNode.IsExpanded)
                    {
                        childNode.Expand();
                    }
                    
                    // Continue recursively
                    ExpandToTargetPath(childNode, targetPath);
                    return;
                }
            }
        }

        private void ListView_BeforeLabelEdit(object sender, LabelEditEventArgs e)
        {
            // Allow editing only if we're not in special views
            if (string.IsNullOrEmpty(navigationManager.CurrentPath))
            {
                e.CancelEdit = true;
                return;
            }
            
            // Get the item being edited
            var item = listView.Items[e.Item];
            string itemPath = item.Tag?.ToString() ?? "";
            
            // Don't allow renaming if the file/folder doesn't exist or is not accessible
            if (string.IsNullOrEmpty(itemPath) || (!File.Exists(itemPath) && !Directory.Exists(itemPath)))
            {
                e.CancelEdit = true;
                return;
            }
            
            Debug.WriteLine($"Starting rename for: {itemPath}");
        }

        private void ListView_AfterLabelEdit(object sender, LabelEditEventArgs e)
        {
            // If editing was cancelled or label is empty, don't do anything
            if (e.Label == null || string.IsNullOrWhiteSpace(e.Label))
            {
                e.CancelEdit = true;
                return;
            }
            
            try
            {
                var item = listView.Items[e.Item];
                string oldPath = item.Tag?.ToString() ?? "";
                string newName = e.Label.Trim();
                
                if (string.IsNullOrEmpty(oldPath))
                {
                    e.CancelEdit = true;
                    return;
                }
                
                // Get the current filename to compare with new name
                string currentFileName = Path.GetFileName(oldPath);
                
                // Check if the new name is the same as the current name (no-op)
                if (string.Equals(currentFileName, newName, StringComparison.OrdinalIgnoreCase))
                {
                    Debug.WriteLine($"Rename cancelled: new name '{newName}' is the same as current name");
                    e.CancelEdit = true;
                    return;
                }
                
                // Validate the new name
                if (!IsValidFileName(newName))
                {
                    MessageBox.Show("The filename contains invalid characters.\n\nA filename cannot contain any of the following characters:\n\\ / : * ? \" < > |", 
                        "Rename Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.CancelEdit = true;
                    return;
                }
                
                string parentDirectory = Path.GetDirectoryName(oldPath) ?? "";
                string newPath = Path.Combine(parentDirectory, newName);
                
                // Check if a file/folder with the new name already exists
                if (File.Exists(newPath) || Directory.Exists(newPath))
                {
                    MessageBox.Show($"A file or folder with the name '{newName}' already exists.", 
                        "Rename Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.CancelEdit = true;
                    return;
                }
                
                // Perform the rename operation
                bool success = false;
                try
                {
                    if (Directory.Exists(oldPath))
                    {
                        Directory.Move(oldPath, newPath);
                        success = true;
                        Debug.WriteLine($"Renamed directory: {oldPath} -> {newPath}");
                    }
                    else if (File.Exists(oldPath))
                    {
                        File.Move(oldPath, newPath);
                        success = true;
                        Debug.WriteLine($"Renamed file: {oldPath} -> {newPath}");
                    }
                    
                    if (success)
                    {
                        // Update the item's tag with the new path
                        item.Tag = newPath;
                        
                        // Refresh the current directory to update any changes
                        this.BeginInvoke(new Action(() => {
                            BtnRefresh_Click(this, EventArgs.Empty);
                            
                            // Try to select the renamed item
                            foreach (ListViewItem listItem in listView.Items)
                            {
                                if (listItem.Tag?.ToString() == newPath)
                                {
                                    listItem.Selected = true;
                                    listItem.Focused = true;
                                    listItem.EnsureVisible();
                                    break;
                                }
                            }
                        }));
                    }
                    else
                    {
                        e.CancelEdit = true;
                    }
                }
                catch (UnauthorizedAccessException)
                {
                    MessageBox.Show("Access denied. You don't have permission to rename this item.", 
                        "Rename Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.CancelEdit = true;
                }
                catch (DirectoryNotFoundException)
                {
                    MessageBox.Show("The directory path was not found.", 
                        "Rename Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.CancelEdit = true;
                }
                catch (IOException ex)
                {
                    MessageBox.Show($"An error occurred while renaming:\n{ex.Message}", 
                        "Rename Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.CancelEdit = true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Unexpected error while renaming:\n{ex.Message}", 
                        "Rename Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.CancelEdit = true;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error in ListView_AfterLabelEdit: {ex.Message}");
                MessageBox.Show($"An error occurred during rename: {ex.Message}", 
                    "Rename Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.CancelEdit = true;
            }
        }

        private bool IsValidFileName(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName))
                return false;
                
            // Check for invalid characters
            char[] invalidChars = Path.GetInvalidFileNameChars();
            foreach (char c in invalidChars)
            {
                if (fileName.Contains(c))
                    return false;
            }
            
            // Check for reserved names on Windows
            string[] reservedNames = { "CON", "PRN", "AUX", "NUL", 
                "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", 
                "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9" };
            
            string nameWithoutExtension = Path.GetFileNameWithoutExtension(fileName).ToUpperInvariant();
            if (reservedNames.Contains(nameWithoutExtension))
                return false;
            
            // Check if filename ends with a dot or space (not allowed on Windows)
            if (fileName.EndsWith(".") || fileName.EndsWith(" "))
                return false;
                
            return true;
        }

        private void ResetAllSettings()
        {
            var result = MessageBox.Show(
                "This will reset all application settings including:\n\n" +
                "• Window size and position\n" +
                "• Toolbar positions\n" +
                "• Font settings\n" +
                "• All favorites/bookmarks\n\n" +
                "Are you sure you want to continue?",
                "Reset All Settings",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning
            );
            
            if (result == DialogResult.Yes)
            {
                try
                {
                    // Clear all registry settings
                    registryManager.ClearAllSettings();
                    
                    MessageBox.Show(
                        "All settings have been reset. The application will now restart to apply the default settings.",
                        "Settings Reset",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                    
                    // Restart the application
                    Application.Restart();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        $"Error resetting settings: {ex.Message}",
                        "Reset Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                }
            }
        }

        private void ExportSettings()
        {
            try
            {
                using (var saveDialog = new SaveFileDialog())
                {
                    saveDialog.Filter = "Registry Files (*.reg)|*.reg|All Files (*.*)|*.*";
                    saveDialog.DefaultExt = "reg";
                    saveDialog.FileName = "win9xplorer_settings.reg";
                    saveDialog.Title = "Export Settings";
                    
                    if (saveDialog.ShowDialog(this) == DialogResult.OK)
                    {
                        // Export registry key to .reg file
                        string regPath = @"HKEY_CURRENT_USER\SOFTWARE\win9xplorer";
                        string regCommand = $"regedit.exe /e \"{saveDialog.FileName}\" \"{regPath}\"";
                        
                        var process = new System.Diagnostics.Process();
                        process.StartInfo.FileName = "cmd.exe";
                        process.StartInfo.Arguments = $"/c {regCommand}";
                        process.StartInfo.UseShellExecute = false;
                        process.StartInfo.CreateNoWindow = true;
                        process.Start();
                        process.WaitForExit();
                        
                        if (process.ExitCode == 0)
                        {
                            MessageBox.Show($"Settings exported successfully to:\n{saveDialog.FileName}", 
                                "Export Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            MessageBox.Show("Failed to export settings.", 
                                "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error exporting settings: {ex.Message}", 
                    "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        private void ImportSettings()
        {
            try
            {
                using (var openDialog = new OpenFileDialog())
                {
                    openDialog.Filter = "Registry Files (*.reg)|*.reg|All Files (*.*)|*.*";
                    openDialog.Title = "Import Settings";
                    
                    if (openDialog.ShowDialog(this) == DialogResult.OK)
                    {
                        var result = MessageBox.Show(
                            "This will overwrite your current application settings.\n\n" +
                            "Are you sure you want to import settings from:\n" +
                            openDialog.FileName,
                            "Import Settings",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Question
                        );
                        
                        if (result == DialogResult.Yes)
                        {
                            // Import .reg file
                            var process = new System.Diagnostics.Process();
                            process.StartInfo.FileName = "regedit.exe";
                            process.StartInfo.Arguments = $"/s \"{openDialog.FileName}\"";
                            process.StartInfo.UseShellExecute = false;
                            process.StartInfo.CreateNoWindow = true;
                            process.Start();
                            process.WaitForExit();
                            
                            if (process.ExitCode == 0)
                            {
                                MessageBox.Show(
                                    "Settings imported successfully. The application will now restart to apply the imported settings.",
                                    "Import Complete",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Information
                                );
                                
                                Application.Restart();
                            }
                            else
                            {
                                MessageBox.Show("Failed to import settings.", 
                                    "Import Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error importing settings: {ex.Message}", 
                    "Import Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        
        private void BtnBack_Click(object sender, EventArgs e)
        {
            var path = navigationManager.GoBack();
            if (path != null)
            {
                suggestionManager.SetNavigatingProgrammatically(true);
                
                navigationManager.CurrentPath = path;
                txtAddress.Text = path;
                
                suggestionManager.SetNavigatingProgrammatically(false);
                
                fileSystemManager.LoadDirectoryContents(path, listView);
                UpdateNavigationButtons();
                UpdateStatusBar();
                UpdateWindowTitle();
                UpdateBookmarkButtonState();
                SaveCurrentPath();
            }
        }

        private void BtnForward_Click(object sender, EventArgs e)
        {
            var path = navigationManager.GoForward();
            if (path != null)
            {
                suggestionManager.SetNavigatingProgrammatically(true);
                
                navigationManager.CurrentPath = path;
                txtAddress.Text = path;
                
                suggestionManager.SetNavigatingProgrammatically(false);
                
                fileSystemManager.LoadDirectoryContents(path, listView);
                UpdateNavigationButtons();
                UpdateStatusBar();
                UpdateWindowTitle();
                UpdateBookmarkButtonState();
                SaveCurrentPath();
            }
        }

        private void BtnUp_Click(object sender, EventArgs e)
        {
            var parentPath = navigationManager.GoUp();
            if (parentPath != null)
            {
                if (parentPath == "")
                {
                    // Go to My Computer
                    NavigateToMyComputer();
                }
                else
                {
                    NavigateToPath(parentPath);
                }
            }
        }

        private void UpdateWindowTitle()
        {
            // Set window title to match the address bar
            string title = txtAddress.Text;
            if (string.IsNullOrEmpty(title) || title == "My Computer")
            {
                this.Text = "My Computer";
            }
            else if (title == "Favorites")
            {
                this.Text = "Favorites";
            }
            else
            {
                // For regular paths, show just the folder name in the title
                try
                {
                    var directoryInfo = new DirectoryInfo(title);
                    this.Text = directoryInfo.Name;
                }
                catch
                {
                    this.Text = title;
                }
            }
        }

        private void UpdateNavigationButtons()
        {
            // Special case for My Computer and Favorites view (when currentPath is empty)
            if (string.IsNullOrEmpty(navigationManager.CurrentPath))
            {
                btnBack.Enabled = false;
                btnForward.Enabled = false;
                btnUp.Enabled = false;
                return;
            }
            
            btnBack.Enabled = navigationManager.CanGoBack;
            btnForward.Enabled = navigationManager.CanGoForward;
            btnUp.Enabled = navigationManager.CanGoUp;
        }

        private void UpdateStatusBar()
        {
            int itemCount = listView.Items.Count;
            int folderCount = listView.Items.Cast<ListViewItem>()
                .Count(item => Directory.Exists(item.Tag?.ToString() ?? ""));
            int fileCount = itemCount - folderCount;
            
            statusLabel.Text = $"{itemCount} object(s) ({folderCount} folder(s), {fileCount} file(s))";
        }

        private void ListView_MouseUp(object sender, MouseEventArgs e)
        {
            Debug.WriteLine($"ListView_MouseUp called - Button: {e.Button}, Location: {e.Location}");
            
            // Only handle right-click up events for context menu
            if (e.Button == MouseButtons.Right)
            {
                contextMenuManager.ShowContextMenu(listView, e, navigationManager.CurrentPath, 
                    () => BtnRefresh_Click(this, EventArgs.Empty), 
                    (view) => SetViewMode(view));
            }
        }

        private void ListView_ItemActivate(object sender, EventArgs e)
        {
            if (listView.SelectedItems.Count > 0)
            {
                var item = listView.SelectedItems[0];
                string path = item.Tag?.ToString() ?? "";

                // Handle drives from My Computer view
                if (string.IsNullOrEmpty(navigationManager.CurrentPath) && !string.IsNullOrEmpty(path))
                {
                    if (txtAddress.Text == "Favorites")
                    {
                        // This is a bookmark from Favorites view - don't sync tree view
                        if (path == "My Computer")
                        {
                            NavigateToMyComputer();
                        }
                        else if (Directory.Exists(path))
                        {
                            NavigateToPath(path, false); // Don't sync tree view for favorites navigation
                        }
                    }
                    else
                    {
                        // This is a drive from My Computer view - sync tree view normally
                        if (Directory.Exists(path))
                        {
                            NavigateToPath(path, true);
                        }
                    }
                    return;
                }

                // Handle regular files and folders - sync tree view normally
                if (Directory.Exists(path))
                {
                    NavigateToPath(path, true);
                }
                else if (File.Exists(path))
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
                        MessageBox.Show($"Error opening file: {ex.Message}", "Error", 
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
        }

        private void BtnRefresh_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(navigationManager.CurrentPath))
            {
                if (txtAddress.Text == "My Computer")
                {
                    // Refresh My Computer view
                    LoadMyComputerContents();
                }
                else if (txtAddress.Text == "Favorites")
                {
                    // Refresh Favorites view
                    LoadFavoritesContents();
                    RefreshFavoritesTreeView();
                }
            }
            else if (!string.IsNullOrEmpty(navigationManager.CurrentPath))
            {
                // Refresh regular directory
                fileSystemManager.LoadDirectoryContents(navigationManager.CurrentPath, listView);
            }
            
            UpdateStatusBar();
        }

        private void TxtAddress_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                string path = txtAddress.Text;
                bool navigationSuccessful = false;
                
                if (Directory.Exists(path))
                {
                    NavigateToPath(path);
                    navigationSuccessful = true;
                }
                else if (path.Equals("My Computer", StringComparison.OrdinalIgnoreCase))
                {
                    NavigateToMyComputer();
                    navigationSuccessful = true;
                }
                else if (path.Equals("Favorites", StringComparison.OrdinalIgnoreCase))
                {
                    // Navigate to Favorites
                    foreach (TreeNode node in treeView.Nodes)
                    {
                        if (node.Tag?.ToString() == "favorites")
                        {
                            treeView.SelectedNode = node;
                            navigationSuccessful = true;
                            break;
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Path not found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    // Restore previous address
                    if (string.IsNullOrEmpty(navigationManager.CurrentPath))
                    {
                        txtAddress.Text = "My Computer";
                    }
                    else
                    {
                        txtAddress.Text = navigationManager.CurrentPath;
                    }
                }
                
                // Move focus to ListView after successful navigation (traditional Windows Explorer behavior)
                if (navigationSuccessful)
                {
                    // Use BeginInvoke to ensure focus is set after all UI updates are complete
                    this.BeginInvoke(new Action(() => {
                        listView.Focus();
                        
                        // If there are items in the ListView, select the first one for immediate keyboard navigation
                        if (listView.Items.Count > 0 && listView.SelectedItems.Count == 0)
                        {
                            listView.Items[0].Selected = true;
                            listView.Items[0].Focused = true;
                        }
                    }));
                }
                
                e.Handled = true;
            }
        }
    }
}
