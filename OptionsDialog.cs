using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace win9xplorer
{
    internal partial class OptionsDialog : Form
    {
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public Font TreeViewFont { get; set; }

        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public Font ListViewFont { get; set; }

        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public FileConflictStrategy ConflictStrategy { get; set; }
        
        private Font originalTreeViewFont;
        private Font originalListViewFont;
        
        public OptionsDialog(Font currentTreeViewFont, Font currentListViewFont, FileConflictStrategy conflictStrategy)
        {
            InitializeComponent();
            
            // Store original fonts for preview and cancel functionality
            originalTreeViewFont = currentTreeViewFont;
            originalListViewFont = currentListViewFont;
            
            // Set current fonts
            TreeViewFont = new Font(currentTreeViewFont, currentTreeViewFont.Style);
            ListViewFont = new Font(currentListViewFont, currentListViewFont.Style);
            ConflictStrategy = conflictStrategy;

            InitializeConflictBehaviorOptions();
            
            // Update preview labels
            UpdateFontPreviews();
            
            // Set classic Windows style
            SetClassicWindowsStyle();
        }

        private void InitializeConflictBehaviorOptions()
        {
            cmbConflictBehavior.Items.Clear();
            cmbConflictBehavior.Items.Add(new ConflictStrategyOption("Ask every time", FileConflictStrategy.AskUser));
            cmbConflictBehavior.Items.Add(new ConflictStrategyOption("Overwrite existing", FileConflictStrategy.OverwriteExisting));
            cmbConflictBehavior.Items.Add(new ConflictStrategyOption("Skip existing", FileConflictStrategy.SkipExisting));

            for (int i = 0; i < cmbConflictBehavior.Items.Count; i++)
            {
                if (cmbConflictBehavior.Items[i] is ConflictStrategyOption option && option.Strategy == ConflictStrategy)
                {
                    cmbConflictBehavior.SelectedIndex = i;
                    return;
                }
            }

            cmbConflictBehavior.SelectedIndex = 0;
        }
        
        private void SetClassicWindowsStyle()
        {
            this.BackColor = SystemColors.Control;
            this.Font = SystemFonts.DefaultFont;
        }
        
        private void UpdateFontPreviews()
        {
            lblTreeViewPreview.Font = TreeViewFont;
            lblTreeViewPreview.Text = $"{TreeViewFont.Name}, {TreeViewFont.SizeInPoints}pt";
            
            lblListViewPreview.Font = ListViewFont;
            lblListViewPreview.Text = $"{ListViewFont.Name}, {ListViewFont.SizeInPoints}pt";
        }
        
        private void BtnTreeViewFont_Click(object sender, EventArgs e)
        {
            using (var fontDialog = new FontDialog())
            {
                fontDialog.Font = TreeViewFont;
                fontDialog.ShowColor = false;
                fontDialog.ShowApply = false;
                fontDialog.ShowEffects = true;
                fontDialog.AllowScriptChange = false;
                
                if (fontDialog.ShowDialog(this) == DialogResult.OK)
                {
                    TreeViewFont?.Dispose();
                    TreeViewFont = new Font(fontDialog.Font, fontDialog.Font.Style);
                    UpdateFontPreviews();
                }
            }
        }
        
        private void BtnListViewFont_Click(object sender, EventArgs e)
        {
            using (var fontDialog = new FontDialog())
            {
                fontDialog.Font = ListViewFont;
                fontDialog.ShowColor = false;
                fontDialog.ShowApply = false;
                fontDialog.ShowEffects = true;
                fontDialog.AllowScriptChange = false;
                
                if (fontDialog.ShowDialog(this) == DialogResult.OK)
                {
                    ListViewFont?.Dispose();
                    ListViewFont = new Font(fontDialog.Font, fontDialog.Font.Style);
                    UpdateFontPreviews();
                }
            }
        }
        
        private void BtnResetTreeViewFont_Click(object sender, EventArgs e)
        {
            TreeViewFont?.Dispose();
            TreeViewFont = new Font(SystemFonts.DefaultFont, SystemFonts.DefaultFont.Style);
            UpdateFontPreviews();
        }
        
        private void BtnResetListViewFont_Click(object sender, EventArgs e)
        {
            ListViewFont?.Dispose();
            ListViewFont = new Font(SystemFonts.DefaultFont, SystemFonts.DefaultFont.Style);
            UpdateFontPreviews();
        }
        
        private void BtnOK_Click(object sender, EventArgs e)
        {
            if (cmbConflictBehavior.SelectedItem is ConflictStrategyOption option)
            {
                ConflictStrategy = option.Strategy;
            }

            this.DialogResult = DialogResult.OK;
            this.Close();
        }
        
        private void BtnCancel_Click(object sender, EventArgs e)
        {
            // Restore original fonts
            TreeViewFont?.Dispose();
            ListViewFont?.Dispose();
            TreeViewFont = new Font(originalTreeViewFont, originalTreeViewFont.Style);
            ListViewFont = new Font(originalListViewFont, originalListViewFont.Style);
            
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
        
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                components?.Dispose();
                TreeViewFont?.Dispose();
                ListViewFont?.Dispose();
            }
            base.Dispose(disposing);
        }

        private sealed class ConflictStrategyOption
        {
            public string DisplayText { get; }
            public FileConflictStrategy Strategy { get; }

            public ConflictStrategyOption(string displayText, FileConflictStrategy strategy)
            {
                DisplayText = displayText;
                Strategy = strategy;
            }

            public override string ToString()
            {
                return DisplayText;
            }
        }
    }
}
