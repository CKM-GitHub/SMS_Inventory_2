namespace SMS.Inventory
{
    partial class frmPETC0306I
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
            this.txtfilepath = new CKM_Controls.CKM_TextBox();
            this.btnForderOpen = new SMS.CustomControls.SMS_Button();
            this.panelBody.SuspendLayout();
            this.panelTitleLeft.SuspendLayout();
            this.panelTitle.SuspendLayout();
            this.SuspendLayout();
            // 
            // panelBody
            // 
            this.panelBody.Controls.Add(this.btnForderOpen);
            this.panelBody.Controls.Add(this.txtfilepath);
            this.panelBody.Controls.Add(this.smS_Label2);
            //this.panelBody.Paint += new System.Windows.Forms.PaintEventHandler(this.panelBody_Paint);
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
            // txtfilepath
            // 
            this.txtfilepath.CKM_Reqired = false;
            this.txtfilepath.CKM_Type = CKM_Controls.CKM_TextBox.Type.Normal;
            this.txtfilepath.DecimalPlace = 0;
            this.txtfilepath.Location = new System.Drawing.Point(135, 16);
            this.txtfilepath.Name = "txtfilepath";
            this.txtfilepath.Size = new System.Drawing.Size(500, 19);
            this.txtfilepath.TabIndex = 1;
            this.txtfilepath.UseMinus = true;
            this.txtfilepath.UseThousandSeparator = true;
            // 
            // btnForderOpen
            // 
            this.btnForderOpen.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.btnForderOpen.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnForderOpen.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnForderOpen.Location = new System.Drawing.Point(635, 16);
            this.btnForderOpen.Name = "btnForderOpen";
            this.btnForderOpen.Size = new System.Drawing.Size(30, 19);
            this.btnForderOpen.TabIndex = 2;
            this.btnForderOpen.Text = "▼";
            this.btnForderOpen.UseVisualStyleBackColor = false;
            this.btnForderOpen.Click += new System.EventHandler(this.btnForderOpen_Click);
            // 
            // frmPETC0306I
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1884, 951);
            this.KeyPreview = true;
            this.Location = new System.Drawing.Point(0, 0);
            this.Name = "frmPETC0306I";
            this.Text = "PETC0306I";
            this.Load += new System.EventHandler(this.frmPETC0306I_Load);
            this.panelBody.ResumeLayout(false);
            this.panelBody.PerformLayout();
            this.panelTitleLeft.ResumeLayout(false);
            this.panelTitle.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private CustomControls.SMS_Label smS_Label2;
        private CustomControls.SMS_Button btnForderOpen;
        private CKM_Controls.CKM_TextBox txtfilepath;

    }
}