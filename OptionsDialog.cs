using System;
using System.Drawing;
using System.Windows.Forms;

namespace win9xplorer
{
    public partial class OptionsDialog : Form
    {
        public Font TreeViewFont { get; set; }
        public Font ListViewFont { get; set; }
        
        private Font originalTreeViewFont;
        private Font originalListViewFont;
        
        public OptionsDialog(Font currentTreeViewFont, Font currentListViewFont)
        {
            InitializeComponent();
            
            // Store original fonts for preview and cancel functionality
            originalTreeViewFont = currentTreeViewFont;
            originalListViewFont = currentListViewFont;
            
            // Set current fonts
            TreeViewFont = new Font(currentTreeViewFont, currentTreeViewFont.Style);
            ListViewFont = new Font(currentListViewFont, currentListViewFont.Style);
            
            // Update preview labels
            UpdateFontPreviews();
            
            // Set classic Windows style
            SetClassicWindowsStyle();
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
                TreeViewFont?.Dispose();
                ListViewFont?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}