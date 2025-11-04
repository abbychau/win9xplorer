namespace win9xplorer
{
    partial class OptionsDialog
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            // Initialize controls
            grpFonts = new GroupBox();
            lblTreeViewFont = new Label();
            lblTreeViewPreview = new Label();
            btnTreeViewFont = new Button();
            btnResetTreeViewFont = new Button();
            
            lblListViewFont = new Label();
            lblListViewPreview = new Label();
            btnListViewFont = new Button();
            btnResetListViewFont = new Button();
            
            btnOK = new Button();
            btnCancel = new Button();
            
            grpFonts.SuspendLayout();
            SuspendLayout();
            
            // 
            // grpFonts
            // 
            grpFonts.Controls.Add(lblTreeViewFont);
            grpFonts.Controls.Add(lblTreeViewPreview);
            grpFonts.Controls.Add(btnTreeViewFont);
            grpFonts.Controls.Add(btnResetTreeViewFont);
            grpFonts.Controls.Add(lblListViewFont);
            grpFonts.Controls.Add(lblListViewPreview);
            grpFonts.Controls.Add(btnListViewFont);
            grpFonts.Controls.Add(btnResetListViewFont);
            grpFonts.Location = new Point(12, 12);
            grpFonts.Name = "grpFonts";
            grpFonts.Size = new Size(460, 160);
            grpFonts.TabIndex = 0;
            grpFonts.TabStop = false;
            grpFonts.Text = "Fonts";
            
            // 
            // lblTreeViewFont
            // 
            lblTreeViewFont.AutoSize = true;
            lblTreeViewFont.Location = new Point(15, 30);
            lblTreeViewFont.Name = "lblTreeViewFont";
            lblTreeViewFont.Size = new Size(86, 15);
            lblTreeViewFont.TabIndex = 0;
            lblTreeViewFont.Text = "Folder Tree Font:";
            
            // 
            // lblTreeViewPreview
            // 
            lblTreeViewPreview.BorderStyle = BorderStyle.Fixed3D;
            lblTreeViewPreview.Location = new Point(120, 25);
            lblTreeViewPreview.Name = "lblTreeViewPreview";
            lblTreeViewPreview.Size = new Size(200, 25);
            lblTreeViewPreview.TabIndex = 1;
            lblTreeViewPreview.Text = "Sample Font Text";
            lblTreeViewPreview.TextAlign = ContentAlignment.MiddleLeft;
            lblTreeViewPreview.BackColor = Color.White;
            
            // 
            // btnTreeViewFont
            // 
            btnTreeViewFont.Location = new Point(330, 24);
            btnTreeViewFont.Name = "btnTreeViewFont";
            btnTreeViewFont.Size = new Size(60, 27);
            btnTreeViewFont.TabIndex = 2;
            btnTreeViewFont.Text = "Change...";
            btnTreeViewFont.UseVisualStyleBackColor = true;
            btnTreeViewFont.Click += BtnTreeViewFont_Click;
            
            // 
            // btnResetTreeViewFont
            // 
            btnResetTreeViewFont.Location = new Point(395, 24);
            btnResetTreeViewFont.Name = "btnResetTreeViewFont";
            btnResetTreeViewFont.Size = new Size(50, 27);
            btnResetTreeViewFont.TabIndex = 3;
            btnResetTreeViewFont.Text = "Reset";
            btnResetTreeViewFont.UseVisualStyleBackColor = true;
            btnResetTreeViewFont.Click += BtnResetTreeViewFont_Click;
            
            // 
            // lblListViewFont
            // 
            lblListViewFont.AutoSize = true;
            lblListViewFont.Location = new Point(15, 80);
            lblListViewFont.Name = "lblListViewFont";
            lblListViewFont.Size = new Size(78, 15);
            lblListViewFont.TabIndex = 4;
            lblListViewFont.Text = "File List Font:";
            
            // 
            // lblListViewPreview
            // 
            lblListViewPreview.BorderStyle = BorderStyle.Fixed3D;
            lblListViewPreview.Location = new Point(120, 75);
            lblListViewPreview.Name = "lblListViewPreview";
            lblListViewPreview.Size = new Size(200, 25);
            lblListViewPreview.TabIndex = 5;
            lblListViewPreview.Text = "Sample Font Text";
            lblListViewPreview.TextAlign = ContentAlignment.MiddleLeft;
            lblListViewPreview.BackColor = Color.White;
            
            // 
            // btnListViewFont
            // 
            btnListViewFont.Location = new Point(330, 74);
            btnListViewFont.Name = "btnListViewFont";
            btnListViewFont.Size = new Size(60, 27);
            btnListViewFont.TabIndex = 6;
            btnListViewFont.Text = "Change...";
            btnListViewFont.UseVisualStyleBackColor = true;
            btnListViewFont.Click += BtnListViewFont_Click;
            
            // 
            // btnResetListViewFont
            // 
            btnResetListViewFont.Location = new Point(395, 74);
            btnResetListViewFont.Name = "btnResetListViewFont";
            btnResetListViewFont.Size = new Size(50, 27);
            btnResetListViewFont.TabIndex = 7;
            btnResetListViewFont.Text = "Reset";
            btnResetListViewFont.UseVisualStyleBackColor = true;
            btnResetListViewFont.Click += BtnResetListViewFont_Click;
            
            // 
            // btnOK
            // 
            btnOK.Location = new Point(316, 190);
            btnOK.Name = "btnOK";
            btnOK.Size = new Size(75, 30);
            btnOK.TabIndex = 8;
            btnOK.Text = "OK";
            btnOK.UseVisualStyleBackColor = true;
            btnOK.Click += BtnOK_Click;
            
            // 
            // btnCancel
            // 
            btnCancel.DialogResult = DialogResult.Cancel;
            btnCancel.Location = new Point(397, 190);
            btnCancel.Name = "btnCancel";
            btnCancel.Size = new Size(75, 30);
            btnCancel.TabIndex = 9;
            btnCancel.Text = "Cancel";
            btnCancel.UseVisualStyleBackColor = true;
            btnCancel.Click += BtnCancel_Click;
            
            // 
            // OptionsDialog
            // 
            AcceptButton = btnOK;
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            CancelButton = btnCancel;
            ClientSize = new Size(484, 235);
            Controls.Add(grpFonts);
            Controls.Add(btnOK);
            Controls.Add(btnCancel);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "OptionsDialog";
            ShowIcon = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            Text = "Options";
            
            grpFonts.ResumeLayout(false);
            grpFonts.PerformLayout();
            ResumeLayout(false);
        }

        #endregion

        private GroupBox grpFonts;
        private Label lblTreeViewFont;
        private Label lblTreeViewPreview;
        private Button btnTreeViewFont;
        private Button btnResetTreeViewFont;
        private Label lblListViewFont;
        private Label lblListViewPreview;
        private Button btnListViewFont;
        private Button btnResetListViewFont;
        private Button btnOK;
        private Button btnCancel;
    }
}