namespace win9xplorer
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();
            menuStrip = new MenuStrip();
            fileMenu = new ToolStripMenuItem();
            editMenu = new ToolStripMenuItem();
            viewMenu = new ToolStripMenuItem();
            goMenu = new ToolStripMenuItem();
            favoritesMenu = new ToolStripMenuItem();
            toolsMenu = new ToolStripMenuItem();
            helpMenu = new ToolStripMenuItem();
            toolStripContainer = new ToolStripContainer();
            splitContainer = new SplitContainer();
            treeView = new TreeView();
            imageListSmall = new ImageList(components);
            listView = new ListView();
            imageListLarge = new ImageList(components);
            toolStrip = new ToolStrip();
            btnBack = new ToolStripButton();
            btnForward = new ToolStripButton();
            btnUp = new ToolStripButton();
            toolStripSeparator1 = new ToolStripSeparator();
            btnRefresh = new ToolStripButton();
            toolStripSeparator2 = new ToolStripSeparator();
            btnViewLargeIcons = new ToolStripButton();
            btnViewSmallIcons = new ToolStripButton();
            btnViewList = new ToolStripButton();
            btnViewDetails = new ToolStripButton();
            toolStripSeparator3 = new ToolStripSeparator();
            btnToggleTreeView = new ToolStripButton();
            btnNavigateToFolder = new ToolStripButton();
            addressStrip = new ToolStrip();
            lblAddress = new ToolStripLabel();
            txtAddress = new ToolStripTextBox();
            btnBookmark = new ToolStripButton();
            statusStrip = new StatusStrip();
            statusLabel = new ToolStripStatusLabel();
            operationProgressBar = new ToolStripProgressBar();
            menuStrip.SuspendLayout();
            toolStripContainer.ContentPanel.SuspendLayout();
            toolStripContainer.TopToolStripPanel.SuspendLayout();
            toolStripContainer.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)splitContainer).BeginInit();
            splitContainer.Panel1.SuspendLayout();
            splitContainer.Panel2.SuspendLayout();
            splitContainer.SuspendLayout();
            toolStrip.SuspendLayout();
            addressStrip.SuspendLayout();
            statusStrip.SuspendLayout();
            SuspendLayout();
            // 
            // menuStrip
            // 
            menuStrip.ImageScalingSize = new Size(20, 20);
            menuStrip.Items.AddRange(new ToolStripItem[] { fileMenu, editMenu, viewMenu, goMenu, favoritesMenu, toolsMenu, helpMenu });
            menuStrip.Location = new Point(0, 0);
            menuStrip.Name = "menuStrip";
            menuStrip.Padding = new Padding(8, 3, 0, 3);
            menuStrip.Size = new Size(871, 30);
            menuStrip.TabIndex = 0;
            menuStrip.Text = "menuStrip1";
            // 
            // fileMenu
            // 
            fileMenu.Name = "fileMenu";
            fileMenu.Size = new Size(47, 24);
            fileMenu.Text = "&File";
            // 
            // editMenu
            // 
            editMenu.Name = "editMenu";
            editMenu.Size = new Size(50, 24);
            editMenu.Text = "&Edit";
            // 
            // viewMenu
            // 
            viewMenu.Name = "viewMenu";
            viewMenu.Size = new Size(57, 24);
            viewMenu.Text = "&View";
            // 
            // goMenu
            // 
            goMenu.Name = "goMenu";
            goMenu.Size = new Size(43, 24);
            goMenu.Text = "&Go";
            // 
            // favoritesMenu
            // 
            favoritesMenu.Name = "favoritesMenu";
            favoritesMenu.Size = new Size(86, 24);
            favoritesMenu.Text = "F&avorites";
            // 
            // toolsMenu
            // 
            toolsMenu.Name = "toolsMenu";
            toolsMenu.Size = new Size(60, 24);
            toolsMenu.Text = "&Tools";
            // 
            // helpMenu
            // 
            helpMenu.Name = "helpMenu";
            helpMenu.Size = new Size(55, 24);
            helpMenu.Text = "&Help";
            // 
            // toolStripContainer
            // 
            // 
            // toolStripContainer.ContentPanel
            // 
            toolStripContainer.ContentPanel.Controls.Add(splitContainer);
            toolStripContainer.ContentPanel.Size = new Size(871, 277);
            toolStripContainer.Dock = DockStyle.Fill;
            toolStripContainer.Location = new Point(0, 30);
            toolStripContainer.Name = "toolStripContainer";
            toolStripContainer.Size = new Size(871, 329);
            toolStripContainer.TabIndex = 1;
            toolStripContainer.Text = "toolStripContainer1";
            // 
            // toolStripContainer.TopToolStripPanel
            // 
            toolStripContainer.TopToolStripPanel.Controls.Add(toolStrip);
            toolStripContainer.TopToolStripPanel.Controls.Add(addressStrip);
            // 
            // splitContainer
            // 
            splitContainer.Dock = DockStyle.Fill;
            splitContainer.Location = new Point(0, 0);
            splitContainer.Margin = new Padding(4);
            splitContainer.Name = "splitContainer";
            // 
            // splitContainer.Panel1
            // 
            splitContainer.Panel1.Controls.Add(treeView);
            // 
            // splitContainer.Panel2
            // 
            splitContainer.Panel2.Controls.Add(listView);
            splitContainer.Size = new Size(871, 277);
            splitContainer.SplitterDistance = 217;
            splitContainer.SplitterWidth = 5;
            splitContainer.TabIndex = 3;
            // 
            // treeView
            // 
            treeView.Dock = DockStyle.Fill;
            treeView.ImageIndex = 0;
            treeView.ImageList = imageListSmall;
            treeView.Location = new Point(0, 0);
            treeView.Margin = new Padding(4);
            treeView.Name = "treeView";
            treeView.SelectedImageIndex = 0;
            treeView.Size = new Size(217, 277);
            treeView.TabIndex = 0;
            treeView.BeforeExpand += TreeView_BeforeExpand;
            treeView.AfterSelect += TreeView_AfterSelect;
            // 
            // imageListSmall
            // 
            imageListSmall.ColorDepth = ColorDepth.Depth32Bit;
            imageListSmall.ImageSize = new Size(16, 16);
            imageListSmall.TransparentColor = Color.Transparent;
            // 
            // listView
            // 
            listView.Dock = DockStyle.Fill;
            listView.FullRowSelect = true;
            listView.GridLines = true;
            listView.LabelEdit = true;
            listView.LargeImageList = imageListLarge;
            listView.Location = new Point(0, 0);
            listView.Margin = new Padding(4);
            listView.Name = "listView";
            listView.Size = new Size(649, 277);
            listView.SmallImageList = imageListSmall;
            listView.TabIndex = 0;
            listView.UseCompatibleStateImageBehavior = false;
            listView.AfterLabelEdit += ListView_AfterLabelEdit;
            listView.BeforeLabelEdit += ListView_BeforeLabelEdit;
            listView.ItemActivate += ListView_ItemActivate;
            // 
            // imageListLarge
            // 
            imageListLarge.ColorDepth = ColorDepth.Depth32Bit;
            imageListLarge.ImageSize = new Size(32, 32);
            imageListLarge.TransparentColor = Color.Transparent;
            // 
            // toolStrip
            // 
            toolStrip.Dock = DockStyle.None;
            toolStrip.ImageScalingSize = new Size(20, 20);
            toolStrip.Items.AddRange(new ToolStripItem[] { btnBack, btnForward, btnUp, toolStripSeparator1, btnRefresh, toolStripSeparator2, btnViewLargeIcons, btnViewSmallIcons, btnViewList, btnViewDetails, toolStripSeparator3, btnToggleTreeView, btnNavigateToFolder });
            toolStrip.Location = new Point(4, 0);
            toolStrip.Name = "toolStrip";
            toolStrip.Size = new Size(321, 25);
            toolStrip.TabIndex = 1;
            toolStrip.Text = "toolStrip1";
            // 
            // btnBack
            // 
            btnBack.DisplayStyle = ToolStripItemDisplayStyle.Image;
            btnBack.ImageTransparentColor = Color.Magenta;
            btnBack.Name = "btnBack";
            btnBack.Size = new Size(29, 22);
            btnBack.Text = "Back";
            btnBack.Click += BtnBack_Click;
            // 
            // btnForward
            // 
            btnForward.DisplayStyle = ToolStripItemDisplayStyle.Image;
            btnForward.ImageTransparentColor = Color.Magenta;
            btnForward.Name = "btnForward";
            btnForward.Size = new Size(29, 22);
            btnForward.Text = "Forward";
            btnForward.Click += BtnForward_Click;
            // 
            // btnUp
            // 
            btnUp.DisplayStyle = ToolStripItemDisplayStyle.Image;
            btnUp.ImageTransparentColor = Color.Magenta;
            btnUp.Name = "btnUp";
            btnUp.Size = new Size(29, 22);
            btnUp.Text = "Up";
            btnUp.Click += BtnUp_Click;
            // 
            // toolStripSeparator1
            // 
            toolStripSeparator1.Name = "toolStripSeparator1";
            toolStripSeparator1.Size = new Size(6, 25);
            // 
            // btnRefresh
            // 
            btnRefresh.DisplayStyle = ToolStripItemDisplayStyle.Image;
            btnRefresh.ImageTransparentColor = Color.Magenta;
            btnRefresh.Name = "btnRefresh";
            btnRefresh.Size = new Size(29, 22);
            btnRefresh.Text = "Refresh";
            btnRefresh.Click += BtnRefresh_Click;
            // 
            // toolStripSeparator2
            // 
            toolStripSeparator2.Name = "toolStripSeparator2";
            toolStripSeparator2.Size = new Size(6, 25);
            // 
            // btnViewLargeIcons
            // 
            btnViewLargeIcons.DisplayStyle = ToolStripItemDisplayStyle.Image;
            btnViewLargeIcons.ImageTransparentColor = Color.Magenta;
            btnViewLargeIcons.Name = "btnViewLargeIcons";
            btnViewLargeIcons.Size = new Size(29, 22);
            btnViewLargeIcons.Text = "Large Icons";
            btnViewLargeIcons.Click += BtnViewLargeIcons_Click;
            // 
            // btnViewSmallIcons
            // 
            btnViewSmallIcons.DisplayStyle = ToolStripItemDisplayStyle.Image;
            btnViewSmallIcons.ImageTransparentColor = Color.Magenta;
            btnViewSmallIcons.Name = "btnViewSmallIcons";
            btnViewSmallIcons.Size = new Size(29, 22);
            btnViewSmallIcons.Text = "Small Icons";
            btnViewSmallIcons.Click += BtnViewSmallIcons_Click;
            // 
            // btnViewList
            // 
            btnViewList.DisplayStyle = ToolStripItemDisplayStyle.Image;
            btnViewList.ImageTransparentColor = Color.Magenta;
            btnViewList.Name = "btnViewList";
            btnViewList.Size = new Size(29, 22);
            btnViewList.Text = "List";
            btnViewList.Click += BtnViewList_Click;
            // 
            // btnViewDetails
            // 
            btnViewDetails.DisplayStyle = ToolStripItemDisplayStyle.Image;
            btnViewDetails.ImageTransparentColor = Color.Magenta;
            btnViewDetails.Name = "btnViewDetails";
            btnViewDetails.Size = new Size(29, 22);
            btnViewDetails.Text = "Details";
            btnViewDetails.Click += BtnViewDetails_Click;
            // 
            // toolStripSeparator3
            // 
            toolStripSeparator3.Name = "toolStripSeparator3";
            toolStripSeparator3.Size = new Size(6, 25);
            // 
            // btnToggleTreeView
            // 
            btnToggleTreeView.Checked = true;
            btnToggleTreeView.CheckState = CheckState.Checked;
            btnToggleTreeView.DisplayStyle = ToolStripItemDisplayStyle.Image;
            btnToggleTreeView.ImageTransparentColor = Color.Magenta;
            btnToggleTreeView.Name = "btnToggleTreeView";
            btnToggleTreeView.Size = new Size(29, 22);
            btnToggleTreeView.Text = "Folders";
            btnToggleTreeView.ToolTipText = "Show/Hide Folder Tree";
            btnToggleTreeView.Click += BtnToggleTreeView_Click;
            // 
            // btnNavigateToFolder
            // 
            btnNavigateToFolder.DisplayStyle = ToolStripItemDisplayStyle.Image;
            btnNavigateToFolder.ImageTransparentColor = Color.Magenta;
            btnNavigateToFolder.Name = "btnNavigateToFolder";
            btnNavigateToFolder.Size = new Size(29, 22);
            btnNavigateToFolder.Text = "Navigate to Folder";
            btnNavigateToFolder.ToolTipText = "Navigate to Folder in Tree";
            btnNavigateToFolder.Click += BtnNavigateToFolder_Click;
            // 
            // addressStrip
            // 
            addressStrip.Dock = DockStyle.None;
            addressStrip.ImageScalingSize = new Size(20, 20);
            addressStrip.Items.AddRange(new ToolStripItem[] { lblAddress, txtAddress, btnBookmark });
            addressStrip.Location = new Point(4, 25);
            addressStrip.Name = "addressStrip";
            addressStrip.Size = new Size(414, 27);
            addressStrip.TabIndex = 2;
            addressStrip.Text = "addressStrip1";
            // 
            // lblAddress
            // 
            lblAddress.Name = "lblAddress";
            lblAddress.Size = new Size(70, 24);
            lblAddress.Text = "Address:";
            // 
            // txtAddress
            // 
            txtAddress.AutoSize = false;
            txtAddress.Name = "txtAddress";
            txtAddress.Size = new Size(300, 27);
            txtAddress.KeyPress += TxtAddress_KeyPress;
            // 
            // btnBookmark
            // 
            btnBookmark.DisplayStyle = ToolStripItemDisplayStyle.Image;
            btnBookmark.ImageTransparentColor = Color.Magenta;
            btnBookmark.Name = "btnBookmark";
            btnBookmark.Size = new Size(29, 24);
            btnBookmark.Text = "Add to Favorites";
            btnBookmark.ToolTipText = "Add to Favorites";
            btnBookmark.Click += BtnBookmark_Click;
            // 
            // statusStrip
            // 
            statusStrip.ImageScalingSize = new Size(20, 20);
            statusStrip.Items.AddRange(new ToolStripItem[] { statusLabel, operationProgressBar });
            statusStrip.Location = new Point(0, 359);
            statusStrip.Name = "statusStrip";
            statusStrip.Padding = new Padding(1, 0, 18, 0);
            statusStrip.Size = new Size(871, 25);
            statusStrip.TabIndex = 4;
            statusStrip.Text = "statusStrip1";
            // 
            // statusLabel
            // 
            statusLabel.Name = "statusLabel";
            statusLabel.Size = new Size(53, 19);
            statusLabel.Text = "Ready";
            // 
            // operationProgressBar
            // 
            operationProgressBar.MarqueeAnimationSpeed = 30;
            operationProgressBar.Name = "operationProgressBar";
            operationProgressBar.Size = new Size(120, 17);
            operationProgressBar.Style = ProgressBarStyle.Marquee;
            operationProgressBar.Visible = false;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(9F, 19F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(871, 384);
            Controls.Add(toolStripContainer);
            Controls.Add(statusStrip);
            Controls.Add(menuStrip);
            KeyPreview = true;
            MainMenuStrip = menuStrip;
            Margin = new Padding(4);
            Name = "Form1";
            Text = "File Explorer";
            FormClosed += Form1_FormClosed;
            Load += Form1_Load;
            KeyDown += Form1_KeyDown;
            Resize += Form1_Resize;
            menuStrip.ResumeLayout(false);
            menuStrip.PerformLayout();
            toolStripContainer.ContentPanel.ResumeLayout(false);
            toolStripContainer.TopToolStripPanel.ResumeLayout(false);
            toolStripContainer.TopToolStripPanel.PerformLayout();
            toolStripContainer.ResumeLayout(false);
            toolStripContainer.PerformLayout();
            splitContainer.Panel1.ResumeLayout(false);
            splitContainer.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)splitContainer).EndInit();
            splitContainer.ResumeLayout(false);
            toolStrip.ResumeLayout(false);
            toolStrip.PerformLayout();
            addressStrip.ResumeLayout(false);
            addressStrip.PerformLayout();
            statusStrip.ResumeLayout(false);
            statusStrip.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private MenuStrip menuStrip;
        private ToolStripMenuItem fileMenu;
        private ToolStripMenuItem editMenu;
        private ToolStripMenuItem viewMenu;
        private ToolStripMenuItem goMenu;
        private ToolStripMenuItem favoritesMenu;
        private ToolStripMenuItem toolsMenu;
        private ToolStripMenuItem helpMenu;
        
        private ToolStripContainer toolStripContainer;
        private ToolStrip toolStrip;
        private ToolStrip addressStrip;
        private ToolStripButton btnBack;
        private ToolStripButton btnForward;
        private ToolStripButton btnUp;
        private ToolStripSeparator toolStripSeparator1;
        private ToolStripButton btnRefresh;
        private ToolStripSeparator toolStripSeparator2;
        private ToolStripButton btnViewLargeIcons;
        private ToolStripButton btnViewSmallIcons;
        private ToolStripButton btnViewList;
        private ToolStripButton btnViewDetails;
        private ToolStripSeparator toolStripSeparator3;
        private ToolStripButton btnToggleTreeView;
        private ToolStripButton btnNavigateToFolder;
        private ToolStripLabel lblAddress;
        private ToolStripTextBox txtAddress;
        private ToolStripButton btnBookmark;
        
        private SplitContainer splitContainer;
        private TreeView treeView;
        private ListView listView;
        
        private ImageList imageListSmall;
        private ImageList imageListLarge;
        
        private StatusStrip statusStrip;
        private ToolStripStatusLabel statusLabel;
        private ToolStripProgressBar operationProgressBar;
    }
}
