namespace SMS.Order
{
    partial class F0019
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
            this.btnShowDialog = new SMS.CustomControls.SMS_Button();
            this.smS_TextBox1 = new SMS.CustomControls.SMS_TextBox();
            this.cboShip = new System.Windows.Forms.ComboBox();
            this.smS_Label2 = new SMS.CustomControls.SMS_Label();
            this.smS_Label3 = new SMS.CustomControls.SMS_Label();
            this.panelBody.SuspendLayout();
            this.panelTitleLeft.SuspendLayout();
            this.panelTitle.SuspendLayout();
            this.SuspendLayout();
            // 
            // panelBody
            // 
            this.panelBody.Controls.Add(this.smS_Label3);
            this.panelBody.Controls.Add(this.cboShip);
            this.panelBody.Controls.Add(this.smS_Label2);
            this.panelBody.Controls.Add(this.btnShowDialog);
            this.panelBody.Controls.Add(this.smS_TextBox1);
            this.panelBody.TabIndex = 0;
            // 
            // btnShowDialog
            // 
            this.btnShowDialog.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.btnShowDialog.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnShowDialog.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnShowDialog.Location = new System.Drawing.Point(592, 52);
            this.btnShowDialog.Name = "btnShowDialog";
            this.btnShowDialog.Size = new System.Drawing.Size(33, 19);
            this.btnShowDialog.TabIndex = 83;
            this.btnShowDialog.Text = "...";
            this.btnShowDialog.UseVisualStyleBackColor = false;
            this.btnShowDialog.Click += new System.EventHandler(this.btnShowDialog_Click);
            // 
            // smS_TextBox1
            // 
            this.smS_TextBox1.Decimalplace = ((byte)(0));
            this.smS_TextBox1.Location = new System.Drawing.Point(93, 52);
            this.smS_TextBox1.Name = "smS_TextBox1";
            this.smS_TextBox1.Size = new System.Drawing.Size(500, 19);
            this.smS_TextBox1.TabIndex = 1;
            this.smS_TextBox1.txt_Type = SMS.CustomControls.SMS_TextBox.Type.Normal;
            this.smS_TextBox1.UseMinus = false;
            this.smS_TextBox1.UseThousandSeparator = true;
            // 
            // cboShip
            // 
            this.cboShip.FormattingEnabled = true;
            this.cboShip.Location = new System.Drawing.Point(93, 22);
            this.cboShip.Name = "cboShip";
            this.cboShip.Size = new System.Drawing.Size(120, 20);
            this.cboShip.TabIndex = 0;
            // 
            // smS_Label2
            // 
            this.smS_Label2.AutoSize = true;
            this.smS_Label2.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label2.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label2.Location = new System.Drawing.Point(33, 26);
            this.smS_Label2.Name = "smS_Label2";
            this.smS_Label2.Size = new System.Drawing.Size(57, 12);
            this.smS_Label2.TabIndex = 84;
            this.smS_Label2.Text = "運送会社";
            // 
            // smS_Label3
            // 
            this.smS_Label3.AutoSize = true;
            this.smS_Label3.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label3.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label3.Location = new System.Drawing.Point(33, 55);
            this.smS_Label3.Name = "smS_Label3";
            this.smS_Label3.Size = new System.Drawing.Size(57, 12);
            this.smS_Label3.TabIndex = 86;
            this.smS_Label3.Text = "ファイル";
            // 
            // F0019
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1884, 951);
            this.F12_Text = "取込(F12)";
            this.F5_Text = "ｷｬﾝｾﾙ(F5)";
            this.KeyPreview = true;
            this.Location = new System.Drawing.Point(0, 0);
            this.Name = "F0019";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "F0019";
            this.Load += new System.EventHandler(this.F0019_Load);
            this.panelBody.ResumeLayout(false);
            this.panelBody.PerformLayout();
            this.panelTitleLeft.ResumeLayout(false);
            this.panelTitle.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private CustomControls.SMS_Button btnShowDialog;
        private CustomControls.SMS_TextBox smS_TextBox1;
        private CustomControls.SMS_Label smS_Label3;
        private System.Windows.Forms.ComboBox cboShip;
        private CustomControls.SMS_Label smS_Label2;
    }
}