using System.ComponentModel;
using System.Diagnostics;

namespace win9xplorer
{
    public partial class Form1 : Form
    {
        private const string InternalDragFormat = "win9xplorer/InternalDrag";
        private readonly NavigationManager navigationManager;
        private readonly FileSystemManager fileSystemManager;
        private readonly IconManager iconManager;
        private readonly ContextMenuManager contextMenuManager;
        private readonly FileOperationsService fileOperationsService;
        private readonly BookmarkManager bookmarkManager;
        private readonly AddressBarSuggestionManager suggestionManager;
        private readonly RegistrySettingsManager registryManager;
        private System.Windows.Forms.Timer? resizeTimer;
        private System.Windows.Forms.Timer? splitterTimer;
        private System.Windows.Forms.Timer? toolStripTimer;
        private CancellationTokenSource? directoryLoadCts;
        private FileConflictStrategy fileConflictStrategy = FileConflictStrategy.AskUser;
        private ListViewItem? currentDropTargetItem;
        private Color currentDropTargetOriginalBackColor = Color.Empty;
        private Color currentDropTargetOriginalForeColor = Color.Empty;
        private TreeNode? currentTreeDropTargetNode;
        private Color currentTreeDropTargetOriginalBackColor = Color.Empty;
        private Color currentTreeDropTargetOriginalForeColor = Color.Empty;
        
        public Form1()
        {
            InitializeComponent();
            
            // Initialize managers
            iconManager = new IconManager(imageListSmall, imageListLarge);
            fileSystemManager = new FileSystemManager(iconManager);
            navigationManager = new NavigationManager();
            fileOperationsService = new FileOperationsService();
            contextMenuManager = new ContextMenuManager(fileOperationsService);
            bookmarkManager = new BookmarkManager();
            suggestionManager = new AddressBarSuggestionManager();
            registryManager = new RegistrySettingsManager();
            
            // Set up bookmark manager tree refresh callback
            bookmarkManager.SetTreeViewRefreshCallback(RefreshFavoritesTreeView);
            
            SetupImageLists();
            SetupContextMenu();
            SetupDragAndDrop();
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

                // Load file operation settings
                var fileOperationSettings = registryManager.LoadFileOperationSettings();
                fileConflictStrategy = fileOperationSettings.ConflictStrategy;
                
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

                // Save file operation settings
                registryManager.SaveFileOperationSettings(fileConflictStrategy);
                
                // Bookmarks are saved automatically by BookmarkManager
                
                Debug.WriteLine("Application settings saved successfully");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error saving application settings: {ex.Message}");
            }
        }

        private CancellationToken BeginDirectoryLoad()
        {
            directoryLoadCts?.Cancel();
            directoryLoadCts?.Dispose();
            directoryLoadCts = new CancellationTokenSource();
            return directoryLoadCts.Token;
        }

        private async Task<bool> LoadDirectoryContentsAsync(string path)
        {
            var cancellationToken = BeginDirectoryLoad();

            try
            {
                var entries = await fileSystemManager.GetDirectoryEntriesAsync(path, cancellationToken);
                if (cancellationToken.IsCancellationRequested)
                    return false;

                listView.BeginUpdate();
                listView.Items.Clear();

                foreach (var entry in entries)
                {
                    var item = new ListViewItem(entry.Name)
                    {
                        ImageKey = iconManager.GetIconKey(entry.FullPath, entry.IsDirectory),
                        Tag = entry.FullPath
                    };

                    item.SubItems.Add(entry.SizeText);
                    item.SubItems.Add(entry.TypeText);
                    item.SubItems.Add(entry.ModifiedText);
                    listView.Items.Add(item);
                }

                return true;
            }
            catch (OperationCanceledException)
            {
                Debug.WriteLine($"Directory load cancelled: {path}");
                return false;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading directory contents: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            finally
            {
                if (listView.IsHandleCreated)
                {
                    listView.EndUpdate();
                }
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
            this.Icon = iconManager.LoadApplicationIcon();
            
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
            var taskbarMenuItem = new ToolStripMenuItem("Show &Taskbar (9x/ME)", null, (s, e) => ShowRetroTaskbar());
            var exportSettingsMenuItem = new ToolStripMenuItem("&Export Settings...", null, (s, e) => ExportSettings());
            var importSettingsMenuItem = new ToolStripMenuItem("&Import Settings...", null, (s, e) => ImportSettings());
            var resetSettingsMenuItem = new ToolStripMenuItem("&Reset All Settings...", null, (s, e) => ResetAllSettings());
            toolsMenu.DropDownItems.AddRange(new ToolStripItem[] {
                optionsMenuItem,
                taskbarMenuItem,
                new ToolStripSeparator(),
                exportSettingsMenuItem,
                importSettingsMenuItem,
                new ToolStripSeparator(),
                resetSettingsMenuItem
            });

            // Setup Help menu
            var aboutMenuItem = new ToolStripMenuItem("&About File Explorer", null, (s, e) => ShowAbout());
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
            contextMenuManager.SetupOperationStatusCallback(SetOperationStatus);
            
            // Add mouse event handlers - MouseDown for cancellation, MouseUp for showing menu
            listView.MouseDown += ListView_MouseDown;
            listView.MouseUp += ListView_MouseUp;
        }

        private void SetupDragAndDrop()
        {
            listView.AllowDrop = true;
            listView.ItemDrag += ListView_ItemDrag;
            listView.DragEnter += ListView_DragEnter;
            listView.DragOver += ListView_DragOver;
            listView.DragLeave += ListView_DragLeave;
            listView.DragDrop += ListView_DragDrop;

            treeView.AllowDrop = true;
            treeView.DragEnter += TreeView_DragEnter;
            treeView.DragOver += TreeView_DragOver;
            treeView.DragLeave += TreeView_DragLeave;
            treeView.DragDrop += TreeView_DragDrop;
        }

        private void ListView_MouseDown(object? sender, MouseEventArgs e)
        {
            Debug.WriteLine($"ListView_MouseDown called - Button: {e.Button}, Location: {e.Location}");
            
            // Let the context menu manager handle right-click down (for cancellation)
            if (e.Button == MouseButtons.Right)
            {
                contextMenuManager.OnListViewMouseDown(e);
            }
        }

        private void ListView_ItemDrag(object? sender, ItemDragEventArgs e)
        {
            if (string.IsNullOrEmpty(navigationManager.CurrentPath) || listView.SelectedItems.Count == 0)
            {
                return;
            }

            var selectedPaths = new List<string>();
            foreach (ListViewItem item in listView.SelectedItems)
            {
                string itemPath = item.Tag?.ToString() ?? string.Empty;
                if (!string.IsNullOrEmpty(itemPath))
                {
                    selectedPaths.Add(itemPath);
                }
            }

            if (selectedPaths.Count == 0)
            {
                return;
            }

            var files = new System.Collections.Specialized.StringCollection();
            files.AddRange(selectedPaths.ToArray());

            var dataObject = new DataObject();
            dataObject.SetFileDropList(files);
            dataObject.SetData(InternalDragFormat, true);

            listView.DoDragDrop(dataObject, DragDropEffects.Copy | DragDropEffects.Move);
        }

        private void ListView_DragEnter(object? sender, DragEventArgs e)
        {
            e.Effect = GetDropEffect(e);
            UpdateDropTargetHighlight(e);
        }

        private void ListView_DragOver(object? sender, DragEventArgs e)
        {
            e.Effect = GetDropEffect(e);
            UpdateDropTargetHighlight(e);
        }

        private void ListView_DragLeave(object? sender, EventArgs e)
        {
            ClearDropTargetHighlight();
        }

        private void TreeView_DragEnter(object? sender, DragEventArgs e)
        {
            e.Effect = GetTreeDropEffect(e);
            UpdateTreeDropTargetHighlight(e);
        }

        private void TreeView_DragOver(object? sender, DragEventArgs e)
        {
            e.Effect = GetTreeDropEffect(e);
            UpdateTreeDropTargetHighlight(e);
        }

        private void TreeView_DragLeave(object? sender, EventArgs e)
        {
            ClearTreeDropTargetHighlight();
        }

        private async void TreeView_DragDrop(object? sender, DragEventArgs e)
        {
            try
            {
                if (e.Data is null)
                {
                    return;
                }

                string? targetDirectory = ResolveTreeDropTargetDirectory(e);
                if (string.IsNullOrEmpty(targetDirectory))
                {
                    return;
                }

                var sourcePaths = GetDraggedFilePaths(e.Data);
                if (sourcePaths.Count == 0)
                {
                    return;
                }

                bool isMove = e.Effect == DragDropEffects.Move;
                SetOperationStatus(isMove ? "Moving dropped items..." : "Copying dropped items...", true);

                FileOperationResult operationResult;
                try
                {
                    operationResult = await fileOperationsService.PasteToDirectoryAsync(
                        sourcePaths,
                        targetDirectory,
                        isMove,
                        fileConflictStrategy);
                }
                finally
                {
                    RestoreStatusText();
                }

                if (operationResult.IsCanceled)
                {
                    MessageBox.Show("Drag-and-drop operation was canceled.", "Operation Canceled",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                if (operationResult.SuccessCount > 0)
                {
                    if (!string.IsNullOrEmpty(navigationManager.CurrentPath)
                        && navigationManager.CurrentPath.Equals(targetDirectory, StringComparison.OrdinalIgnoreCase))
                    {
                        BtnRefresh_Click(this, EventArgs.Empty);
                    }
                }

                if (operationResult.Errors.Count > 0)
                {
                    ShowDragDropErrors(operationResult.Errors);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error during drag-and-drop: {ex.Message}", "Drag-and-Drop Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally
            {
                ClearTreeDropTargetHighlight();
            }
        }

        private async void ListView_DragDrop(object? sender, DragEventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(navigationManager.CurrentPath))
                {
                    return;
                }

                if (e.Data is null)
                {
                    return;
                }

                var sourcePaths = GetDraggedFilePaths(e.Data);
                if (sourcePaths.Count == 0)
                {
                    return;
                }

                string? targetDirectory = ResolveDropTargetDirectory(e);
                if (string.IsNullOrEmpty(targetDirectory))
                {
                    return;
                }

                bool isMove = e.Effect == DragDropEffects.Move;
                SetOperationStatus(isMove ? "Moving dropped items..." : "Copying dropped items...", true);

                FileOperationResult operationResult;
                try
                {
                    operationResult = await fileOperationsService.PasteToDirectoryAsync(
                        sourcePaths,
                        targetDirectory,
                        isMove,
                        fileConflictStrategy);
                }
                finally
                {
                    RestoreStatusText();
                }

                if (operationResult.IsCanceled)
                {
                    MessageBox.Show("Drag-and-drop operation was canceled.", "Operation Canceled",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                if (operationResult.SuccessCount > 0)
                {
                    BtnRefresh_Click(this, EventArgs.Empty);
                }

                if (operationResult.Errors.Count > 0)
                {
                    ShowDragDropErrors(operationResult.Errors);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error during drag-and-drop: {ex.Message}", "Drag-and-Drop Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally
            {
                ClearDropTargetHighlight();
            }
        }

        private DragDropEffects GetDropEffect(DragEventArgs e)
        {
            if (string.IsNullOrEmpty(navigationManager.CurrentPath) || e.Data is null || !e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                return DragDropEffects.None;
            }

            string? targetDirectory = ResolveDropTargetDirectory(e);
            if (string.IsNullOrEmpty(targetDirectory))
            {
                return DragDropEffects.None;
            }

            if ((e.KeyState & 8) == 8)
            {
                return DragDropEffects.Copy;
            }

            if ((e.KeyState & 4) == 4)
            {
                return DragDropEffects.Move;
            }

            var draggedPaths = GetDraggedFilePaths(e.Data);
            if (draggedPaths.Count == 0)
            {
                return DragDropEffects.None;
            }

            return ShouldDefaultToMove(draggedPaths, targetDirectory)
                ? DragDropEffects.Move
                : DragDropEffects.Copy;
        }

        private DragDropEffects GetTreeDropEffect(DragEventArgs e)
        {
            if (e.Data is null || !e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                return DragDropEffects.None;
            }

            string? targetDirectory = ResolveTreeDropTargetDirectory(e);
            if (string.IsNullOrEmpty(targetDirectory))
            {
                return DragDropEffects.None;
            }

            if ((e.KeyState & 8) == 8)
            {
                return DragDropEffects.Copy;
            }

            if ((e.KeyState & 4) == 4)
            {
                return DragDropEffects.Move;
            }

            var draggedPaths = GetDraggedFilePaths(e.Data);
            if (draggedPaths.Count == 0)
            {
                return DragDropEffects.None;
            }

            return ShouldDefaultToMove(draggedPaths, targetDirectory)
                ? DragDropEffects.Move
                : DragDropEffects.Copy;
        }

        private string? ResolveDropTargetDirectory(DragEventArgs e)
        {
            if (string.IsNullOrEmpty(navigationManager.CurrentPath))
            {
                return null;
            }

            ListViewItem? directoryItem = GetDirectoryItemAtPoint(e);
            if (directoryItem?.Tag?.ToString() is string itemPath && Directory.Exists(itemPath))
            {
                return itemPath;
            }

            return navigationManager.CurrentPath;
        }

        private string? ResolveTreeDropTargetDirectory(DragEventArgs e)
        {
            TreeNode? node = GetTreeNodeAtPoint(e);
            if (node?.Tag?.ToString() is string nodePath && Directory.Exists(nodePath))
            {
                return nodePath;
            }

            return null;
        }

        private ListViewItem? GetDirectoryItemAtPoint(DragEventArgs e)
        {
            Point clientPoint = listView.PointToClient(new Point(e.X, e.Y));
            ListViewHitTestInfo hitTest = listView.HitTest(clientPoint);
            if (hitTest.Item?.Tag?.ToString() is string itemPath && Directory.Exists(itemPath))
            {
                return hitTest.Item;
            }

            return null;
        }

        private TreeNode? GetTreeNodeAtPoint(DragEventArgs e)
        {
            Point clientPoint = treeView.PointToClient(new Point(e.X, e.Y));
            return treeView.GetNodeAt(clientPoint);
        }

        private void UpdateDropTargetHighlight(DragEventArgs e)
        {
            if (e.Effect == DragDropEffects.None)
            {
                ClearDropTargetHighlight();
                return;
            }

            ListViewItem? newTarget = GetDirectoryItemAtPoint(e);
            if (ReferenceEquals(currentDropTargetItem, newTarget))
            {
                return;
            }

            ClearDropTargetHighlight();
            if (newTarget == null)
            {
                return;
            }

            currentDropTargetItem = newTarget;
            currentDropTargetOriginalBackColor = newTarget.BackColor;
            currentDropTargetOriginalForeColor = newTarget.ForeColor;
            newTarget.BackColor = SystemColors.Highlight;
            newTarget.ForeColor = SystemColors.HighlightText;
        }

        private void ClearDropTargetHighlight()
        {
            if (currentDropTargetItem == null)
            {
                return;
            }

            currentDropTargetItem.BackColor = currentDropTargetOriginalBackColor;
            currentDropTargetItem.ForeColor = currentDropTargetOriginalForeColor;
            currentDropTargetItem = null;
        }

        private void UpdateTreeDropTargetHighlight(DragEventArgs e)
        {
            if (e.Effect == DragDropEffects.None)
            {
                ClearTreeDropTargetHighlight();
                return;
            }

            TreeNode? newTarget = GetTreeNodeAtPoint(e);
            if (newTarget?.Tag?.ToString() is not string nodePath || !Directory.Exists(nodePath))
            {
                newTarget = null;
            }

            if (ReferenceEquals(currentTreeDropTargetNode, newTarget))
            {
                return;
            }

            ClearTreeDropTargetHighlight();
            if (newTarget == null)
            {
                return;
            }

            currentTreeDropTargetNode = newTarget;
            currentTreeDropTargetOriginalBackColor = newTarget.BackColor;
            currentTreeDropTargetOriginalForeColor = newTarget.ForeColor;
            newTarget.BackColor = SystemColors.Highlight;
            newTarget.ForeColor = SystemColors.HighlightText;
        }

        private void ClearTreeDropTargetHighlight()
        {
            if (currentTreeDropTargetNode == null)
            {
                return;
            }

            currentTreeDropTargetNode.BackColor = currentTreeDropTargetOriginalBackColor;
            currentTreeDropTargetNode.ForeColor = currentTreeDropTargetOriginalForeColor;
            currentTreeDropTargetNode = null;
        }

        private static List<string> GetDraggedFilePaths(IDataObject dataObject)
        {
            if (!dataObject.GetDataPresent(DataFormats.FileDrop))
            {
                return new List<string>();
            }

            if (dataObject.GetData(DataFormats.FileDrop) is not string[] droppedPaths)
            {
                return new List<string>();
            }

            return droppedPaths.Where(path => !string.IsNullOrWhiteSpace(path)).ToList();
        }

        private static bool ShouldDefaultToMove(IEnumerable<string> sourcePaths, string targetDirectory)
        {
            string targetRoot = Path.GetPathRoot(targetDirectory) ?? string.Empty;
            if (string.IsNullOrEmpty(targetRoot))
            {
                return false;
            }

            foreach (string sourcePath in sourcePaths)
            {
                string sourceRoot = Path.GetPathRoot(sourcePath) ?? string.Empty;
                if (!sourceRoot.Equals(targetRoot, StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }
            }

            return true;
        }

        private void ShowDragDropErrors(List<string> errors)
        {
            var significantErrors = errors
                .Where(error => !IsIgnorableDragDropError(error))
                .ToList();

            if (significantErrors.Count == 0)
            {
                return;
            }

            string errorMessage = string.Join("\n", significantErrors.Take(5));
            if (significantErrors.Count > 5)
            {
                errorMessage += $"\n... and {significantErrors.Count - 5} more errors";
            }

            MessageBox.Show($"Some dropped items failed.\n\n{errorMessage}",
                "Drag-and-Drop Errors", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private static bool IsIgnorableDragDropError(string error)
        {
            return error.Contains("target directory is under source directory", StringComparison.OrdinalIgnoreCase)
                || error.Contains("Source and destination are the same", StringComparison.OrdinalIgnoreCase);
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

                fileOperationsService.CopyToClipboard(selectedPaths);
                
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

                fileOperationsService.CutToClipboard(selectedPaths);
                
                MessageBox.Show($"Cut {selectedPaths.Count} item(s) to clipboard.", "Cut", 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error cutting items: {ex.Message}", "Cut Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private async void PasteItems()
        {
            try
            {
                if (!fileOperationsService.TryGetClipboardFileDrop(out var files, out bool isCutOperation))
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
                
                string operation = isCutOperation ? "move" : "copy";
                var result = MessageBox.Show($"Are you sure you want to {operation} {files.Count} item(s) to this location?", 
                    $"{(isCutOperation ? "Move" : "Copy")} Files", 
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                
                if (result == DialogResult.Yes)
                {
                    SetOperationStatus($"{(isCutOperation ? "Moving" : "Copying")} items...", true);
                    FileOperationResult operationResult;
                    try
                    {
                        operationResult = await fileOperationsService.PasteToDirectoryAsync(files, targetDirectory, isCutOperation, fileConflictStrategy);
                    }
                    finally
                    {
                        RestoreStatusText();
                    }

                    if (operationResult.IsCanceled)
                    {
                        MessageBox.Show($"{(isCutOperation ? "Move" : "Copy")} operation was canceled.",
                            "Operation Canceled",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    if (operationResult.SuccessCount > 0)
                    {
                        // Refresh the current view
                        BtnRefresh_Click(this, EventArgs.Empty);
                        
                        MessageBox.Show($"Successfully {operation}ed {operationResult.SuccessCount} item(s).", 
                            $"{(isCutOperation ? "Move" : "Copy")} Complete", 
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }

                    if (operationResult.Errors.Count > 0)
                    {
                        string errorMessage = string.Join("\n", operationResult.Errors.Take(5));
                        if (operationResult.Errors.Count > 5)
                            errorMessage += $"\n... and {operationResult.Errors.Count - 5} more errors";

                        MessageBox.Show($"Some items failed to {operation}.\n\n{errorMessage}",
                            $"{(isCutOperation ? "Move" : "Copy")} Errors",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }

                    if (operationResult.SkippedCount > 0)
                    {
                        MessageBox.Show($"Skipped {operationResult.SkippedCount} existing item(s).",
                            $"{(isCutOperation ? "Move" : "Copy")} Skipped",
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
            btnToggleTreeView.ToolTipText = splitContainer.Panel1Collapsed ? "Show Folder Tree" : "Hide Folder Tree";
            
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
            using (var optionsDialog = new OptionsDialog(treeView.Font, listView.Font, fileConflictStrategy))
            {
                if (optionsDialog.ShowDialog(this) == DialogResult.OK)
                {
                    // Apply the selected fonts
                    treeView.Font = optionsDialog.TreeViewFont;
                    listView.Font = optionsDialog.ListViewFont;
                    
                    // Save font settings to registry
                    registryManager.SaveFontSettings(treeView.Font, listView.Font);
                    fileConflictStrategy = optionsDialog.ConflictStrategy;
                    registryManager.SaveFileOperationSettings(fileConflictStrategy);
                    
                    Debug.WriteLine($"Options saved. Conflict strategy: {fileConflictStrategy}");
                }
            }
        }

        private void ShowAbout()
        {
            const string projectUrl = "https://github.com/abbychau/win9xplorer";

            using var aboutDialog = new Form
            {
                Text = "About File Explorer",
                FormBorderStyle = FormBorderStyle.FixedDialog,
                StartPosition = FormStartPosition.CenterParent,
                MaximizeBox = false,
                MinimizeBox = false,
                ShowInTaskbar = false,
                ClientSize = new Size(360, 150),
                Font = new Font("MS Sans Serif", 8.25f, FontStyle.Regular, GraphicsUnit.Point)
            };

            var titleLabel = new Label
            {
                AutoSize = true,
                Left = 16,
                Top = 16,
                Text = "Windows 9x Explorer Clone"
            };

            var infoLabel = new Label
            {
                AutoSize = true,
                Left = 16,
                Top = 40,
                Text = "Built with .NET 8\nVersion 1.0"
            };

            var linkLabel = new LinkLabel
            {
                AutoSize = true,
                Left = 16,
                Top = 82,
                Text = projectUrl
            };
            linkLabel.LinkClicked += (_, _) =>
            {
                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = projectUrl,
                        UseShellExecute = true
                    });
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Failed to open project URL: {ex.Message}");
                }
            };

            var okButton = new Button
            {
                Text = "OK",
                Width = 80,
                Height = 24,
                Left = 264,
                Top = 112,
                DialogResult = DialogResult.OK
            };

            aboutDialog.Controls.Add(titleLabel);
            aboutDialog.Controls.Add(infoLabel);
            aboutDialog.Controls.Add(linkLabel);
            aboutDialog.Controls.Add(okButton);
            aboutDialog.AcceptButton = okButton;

            aboutDialog.ShowDialog(this);
        }

        public void OpenPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || !Directory.Exists(path))
            {
                return;
            }

            if (InvokeRequired)
            {
                BeginInvoke(new Action(() => NavigateToPath(path)));
                return;
            }

            NavigateToPath(path);
        }

        private void ShowRetroTaskbar()
        {
            var taskbar = RetroTaskbarForm.GetOrCreate();
            if (!taskbar.Visible)
            {
                taskbar.Show();
            }

            taskbar.Activate();
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

            directoryLoadCts?.Cancel();
            directoryLoadCts?.Dispose();
            directoryLoadCts = null;
            
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
            string currentPath = GetPathForFavoritesAction();
            if (string.IsNullOrEmpty(currentPath))
            {
                btnBookmark.Checked = false;
                btnBookmark.ToolTipText = "Add to Favorites";
                return;
            }

            bool isBookmarked = bookmarkManager.IsBookmarked(currentPath);
            
            btnBookmark.Checked = isBookmarked;
            btnBookmark.ToolTipText = isBookmarked ? "Remove from Favorites" : "Add to Favorites";
        }

        private void BtnBookmark_Click(object sender, EventArgs e)
        {
            string currentPath = GetPathForFavoritesAction();
            if (string.IsNullOrEmpty(currentPath))
            {
                MessageBox.Show("No folder is currently selected to add to Favorites.",
                    "Favorites", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            
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
            string currentPath = GetPathForFavoritesAction();
            if (string.IsNullOrEmpty(currentPath))
            {
                MessageBox.Show("No folder is currently selected to add to Favorites.",
                    "Add to Favorites", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (bookmarkManager.IsBookmarked(currentPath))
            {
                MessageBox.Show("This location is already in Favorites.",
                    "Add to Favorites", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            bookmarkManager.AddBookmark(currentPath);
            UpdateBookmarkButtonState();
            RefreshFavoritesMenu();
            MessageBox.Show("Added to Favorites.", "Favorites", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private string GetPathForFavoritesAction()
        {
            if (!string.IsNullOrEmpty(navigationManager.CurrentPath))
            {
                return navigationManager.CurrentPath;
            }

            if (txtAddress.Text == "My Computer")
            {
                if (listView.SelectedItems.Count == 1)
                {
                    return listView.SelectedItems[0].Tag?.ToString() ?? "My Computer";
                }

                return "My Computer";
            }

            return string.Empty;
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
                    _ = LoadDirectoryContentsAsync(navigationManager.CurrentPath);
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

        private async void TreeView_BeforeExpand(object sender, TreeViewCancelEventArgs e)
        {
            if (e.Node?.Tag?.ToString() == "computer" || e.Node?.Tag?.ToString() == "favorites")
                return;

            if (e.Node?.Tag?.ToString() is not string path || string.IsNullOrEmpty(path))
                return;

            try
            {
                using var cts = new CancellationTokenSource();
                var directories = await fileSystemManager.GetTreeDirectoryInfosAsync(path, cts.Token);

                e.Node.Nodes.Clear();
                foreach (var directory in directories)
                {
                    var folderNode = new TreeNode(directory.Name)
                    {
                        ImageKey = "folder",
                        SelectedImageKey = "folder",
                        Tag = directory.FullPath
                    };

                    if (directory.HasChildren)
                    {
                        folderNode.Nodes.Add(new TreeNode("Loading..."));
                    }

                    e.Node.Nodes.Add(folderNode);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading folders: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
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

            View currentView = listView.View;
            listView.Clear();
            listView.View = currentView;
            
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
            _ = NavigateToPathAsync(path, true);
        }

        private void NavigateToPath(string path, bool syncTreeView)
        {
            _ = NavigateToPathAsync(path, syncTreeView);
        }

        private async Task NavigateToPathAsync(string path, bool syncTreeView)
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

                if (!await LoadDirectoryContentsAsync(path))
                {
                    return;
                }
                
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

        private async void DeleteSelectedItems()
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
                    ? $"Move '{Path.GetFileName(selectedPaths[0])}' to the Recycle Bin?"
                    : $"Move these {selectedPaths.Count} items to the Recycle Bin?";
                
                var result = MessageBox.Show(message, "Delete Files", 
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                
                if (result == DialogResult.Yes)
                {
                    SetOperationStatus("Moving items to Recycle Bin...", true);
                    FileOperationResult deleteResult;
                    try
                    {
                        deleteResult = await fileOperationsService.DeletePathsAsync(selectedPaths);
                    }
                    finally
                    {
                        RestoreStatusText();
                    }

                    if (deleteResult.IsCanceled)
                    {
                        MessageBox.Show("Delete operation was canceled.", "Operation Canceled",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                    
                    // Refresh the current view to reflect deletions
                    BtnRefresh_Click(this, EventArgs.Empty);
                    
                    // Show results
                    if (deleteResult.Errors.Count == 0)
                    {
                        Debug.WriteLine($"Successfully deleted {deleteResult.SuccessCount} items");
                    }
                    else if (deleteResult.SuccessCount > 0)
                    {
                        string errorMessage = $"Successfully deleted {deleteResult.SuccessCount} items.\n\nErrors:\n" + 
                                            string.Join("\n", deleteResult.Errors.Take(5));
                        if (deleteResult.Errors.Count > 5)
                            errorMessage += $"\n... and {deleteResult.Errors.Count - 5} more errors";
                             
                        MessageBox.Show(errorMessage, "Delete Partial Success", 
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    else
                    {
                        string errorMessage = "Failed to delete items:\n" + 
                                            string.Join("\n", deleteResult.Errors.Take(5));
                        if (deleteResult.Errors.Count > 5)
                            errorMessage += $"\n... and {deleteResult.Errors.Count - 5} more errors";
                             
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
            EnsureTreeViewVisible();

            if (txtAddress.Text == "My Computer")
            {
                SelectTreeRootNode("computer");
                return;
            }

            if (txtAddress.Text == "Favorites")
            {
                SelectTreeRootNode("favorites");
                return;
            }

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

        private void EnsureTreeViewVisible()
        {
            if (splitContainer.Panel1Collapsed)
            {
                ToggleTreeView();
            }
        }

        private void SelectTreeRootNode(string rootTag)
        {
            foreach (TreeNode node in treeView.Nodes)
            {
                if (string.Equals(node.Tag?.ToString(), rootTag, StringComparison.OrdinalIgnoreCase))
                {
                    treeView.SelectedNode = node;
                    node.EnsureVisible();
                    break;
                }
            }
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
        
        private async void BtnBack_Click(object sender, EventArgs e)
        {
            var path = navigationManager.GoBack();
            if (path != null)
            {
                suggestionManager.SetNavigatingProgrammatically(true);
                
                navigationManager.CurrentPath = path;
                txtAddress.Text = path;
                
                suggestionManager.SetNavigatingProgrammatically(false);
                
                if (!await LoadDirectoryContentsAsync(path))
                {
                    return;
                }
                UpdateNavigationButtons();
                UpdateStatusBar();
                UpdateWindowTitle();
                UpdateBookmarkButtonState();
                SaveCurrentPath();
            }
        }

        private async void BtnForward_Click(object sender, EventArgs e)
        {
            var path = navigationManager.GoForward();
            if (path != null)
            {
                suggestionManager.SetNavigatingProgrammatically(true);
                
                navigationManager.CurrentPath = path;
                txtAddress.Text = path;
                
                suggestionManager.SetNavigatingProgrammatically(false);
                
                if (!await LoadDirectoryContentsAsync(path))
                {
                    return;
                }
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

        private void SetOperationStatus(string message, bool inProgress)
        {
            if (!inProgress)
            {
                RestoreStatusText();
                return;
            }

            statusLabel.Text = message;
            operationProgressBar.Visible = inProgress;
        }

        private void RestoreStatusText()
        {
            operationProgressBar.Visible = false;

            if (string.IsNullOrEmpty(navigationManager.CurrentPath))
            {
                if (txtAddress.Text == "My Computer")
                {
                    statusLabel.Text = $"{listView.Items.Count} drive(s)";
                    return;
                }

                if (txtAddress.Text == "Favorites")
                {
                    statusLabel.Text = $"{listView.Items.Count} favorite(s)";
                    return;
                }
            }

            UpdateStatusBar();
        }

        private void ListView_MouseUp(object? sender, MouseEventArgs e)
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

        private async void BtnRefresh_Click(object sender, EventArgs e)
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
                await LoadDirectoryContentsAsync(navigationManager.CurrentPath);
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
