namespace SMS.Inventory
{
    partial class frmPETC0305I
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.smS_Label2 = new SMS.CustomControls.SMS_Label();
            this.txtpath = new CKM_Controls.CKM_TextBox();
            this.btnFolderOpen = new SMS.CustomControls.SMS_Button();
            this.panelBody.SuspendLayout();
            this.panelTitleLeft.SuspendLayout();
            this.panelTitle.SuspendLayout();
            this.SuspendLayout();
            // 
            // panelBody
            // 
            this.panelBody.Controls.Add(this.btnFolderOpen);
            this.panelBody.Controls.Add(this.txtpath);
            this.panelBody.Controls.Add(this.smS_Label2);
            // 
            // smS_Label2
            // 
            this.smS_Label2.AutoSize = true;
            this.smS_Label2.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label2.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label2.Location = new System.Drawing.Point(46, 21);
            this.smS_Label2.Name = "smS_Label2";
            this.smS_Label2.Size = new System.Drawing.Size(83, 12);
            this.smS_Label2.TabIndex = 0;
            this.smS_Label2.Text = "取込ファイル";
            // 
            // txtpath
            // 
            this.txtpath.CKM_Reqired = false;
            this.txtpath.CKM_Type = CKM_Controls.CKM_TextBox.Type.Normal;
            this.txtpath.DecimalPlace = 0;
            this.txtpath.Location = new System.Drawing.Point(157, 21);
            this.txtpath.Name = "txtpath";
            this.txtpath.Size = new System.Drawing.Size(500, 19);
            this.txtpath.TabIndex = 1;
            this.txtpath.UseMinus = true;
            this.txtpath.UseThousandSeparator = true;
            // 
            // btnFolderOpen
            // 
            this.btnFolderOpen.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.btnFolderOpen.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnFolderOpen.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnFolderOpen.Location = new System.Drawing.Point(653, 21);
            this.btnFolderOpen.Name = "btnFolderOpen";
            this.btnFolderOpen.Size = new System.Drawing.Size(30, 19);
            this.btnFolderOpen.TabIndex = 2;
            this.btnFolderOpen.Text = "▼";
            this.btnFolderOpen.UseVisualStyleBackColor = false;
            this.btnFolderOpen.Click += new System.EventHandler(this.btnFolderOpen_Click);
            // 
            // frmPETC0305I
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1884, 951);
            this.KeyPreview = true;
            this.Location = new System.Drawing.Point(0, 0);
            this.Name = "frmPETC0305I";
            this.Text = "PETC0305I";
            this.Load += new System.EventHandler(this.frmPETC0305I_Load);
            this.panelBody.ResumeLayout(false);
            this.panelBody.PerformLayout();
            this.panelTitleLeft.ResumeLayout(false);
            this.panelTitle.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private CustomControls.SMS_Label smS_Label2;
        private CKM_Controls.CKM_TextBox txtpath;
        private CustomControls.SMS_Button btnFolderOpen;
    }
}