namespace SMS.Login
{
    partial class frmLogin
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmLogin));
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.label1 = new System.Windows.Forms.Label();
            this.lblSProcessingDate = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.lblVersion = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.lblJSYRN = new System.Windows.Forms.Label();
            this.txtKAICD = new SMS.CustomControls.SMS_TextBox();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.panel3 = new System.Windows.Forms.Panel();
            this.txtPASWD = new SMS.CustomControls.SMS_TextBox();
            this.btnChange = new SMS.CustomControls.SMS_Button();
            this.smS_Label1 = new SMS.CustomControls.SMS_Label();
            this.smS_Label2 = new SMS.CustomControls.SMS_Label();
            this.smS_Label3 = new SMS.CustomControls.SMS_Label();
            this.txtTANCD = new SMS.CustomControls.SMS_TextBox();
            this.panel4 = new System.Windows.Forms.Panel();
            this.panel5 = new System.Windows.Forms.Panel();
            this.bnF1 = new SMS.CustomControls.SMS_Button();
            this.bnF12 = new SMS.CustomControls.SMS_Button();
            this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
            this.btnF12 = new SMS.CustomControls.SMS_Button();
            this.btnF1 = new SMS.CustomControls.SMS_Button();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            this.panel3.SuspendLayout();
            this.panel4.SuspendLayout();
            this.panel5.SuspendLayout();
            this.SuspendLayout();
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackColor = System.Drawing.Color.Transparent;
            this.pictureBox1.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("pictureBox1.BackgroundImage")));
            this.pictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.pictureBox1.Location = new System.Drawing.Point(0, 0);
            this.pictureBox1.Margin = new System.Windows.Forms.Padding(4, 2, 4, 2);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(255, 77);
            this.pictureBox1.TabIndex = 43;
            this.pictureBox1.TabStop = false;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(301, 14);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(44, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "処理日";
            // 
            // lblSProcessingDate
            // 
            this.lblSProcessingDate.BackColor = System.Drawing.Color.Transparent;
            this.lblSProcessingDate.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblSProcessingDate.Location = new System.Drawing.Point(348, 8);
            this.lblSProcessingDate.Name = "lblSProcessingDate";
            this.lblSProcessingDate.Size = new System.Drawing.Size(104, 24);
            this.lblSProcessingDate.TabIndex = 1;
            this.lblSProcessingDate.Text = "9999/99/99";
            this.lblSProcessingDate.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.lblVersion);
            this.panel1.Controls.Add(this.pictureBox1);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.lblSProcessingDate);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(465, 79);
            this.panel1.TabIndex = 0;
            // 
            // lblVersion
            // 
            this.lblVersion.AutoSize = true;
            this.lblVersion.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblVersion.ForeColor = System.Drawing.Color.Red;
            this.lblVersion.Location = new System.Drawing.Point(300, 50);
            this.lblVersion.Name = "lblVersion";
            this.lblVersion.Size = new System.Drawing.Size(93, 18);
            this.lblVersion.TabIndex = 44;
            this.lblVersion.Text = "Version 2.0";
            // 
            // panel2
            // 
            this.panel2.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.panel2.Controls.Add(this.lblJSYRN);
            this.panel2.Controls.Add(this.txtKAICD);
            this.panel2.Location = new System.Drawing.Point(110, 9);
            this.panel2.Margin = new System.Windows.Forms.Padding(0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(324, 23);
            this.panel2.TabIndex = 0;
            this.panel2.TabStop = true;
            // 
            // lblJSYRN
            // 
            this.lblJSYRN.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(169)))), ((int)(((byte)(208)))), ((int)(((byte)(142)))));
            this.lblJSYRN.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblJSYRN.Location = new System.Drawing.Point(3, 2);
            this.lblJSYRN.Name = "lblJSYRN";
            this.lblJSYRN.Size = new System.Drawing.Size(240, 19);
            this.lblJSYRN.TabIndex = 44;
            this.lblJSYRN.Text = "CapitalSports";
            this.lblJSYRN.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // txtKAICD
            // 
            this.txtKAICD.Decimalplace = ((byte)(0));
            this.txtKAICD.Enabled = false;
            this.txtKAICD.Location = new System.Drawing.Point(265, 2);
            this.txtKAICD.MaxLength = 2;
            this.txtKAICD.Name = "txtKAICD";
            this.txtKAICD.Size = new System.Drawing.Size(30, 19);
            this.txtKAICD.TabIndex = 1;
            this.txtKAICD.Text = "01";
            this.txtKAICD.txt_Type = SMS.CustomControls.SMS_TextBox.Type.IntegerOnly;
            this.txtKAICD.UseMinus = false;
            this.txtKAICD.UseThousandSeparator = true;
            this.txtKAICD.Visible = false;
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.BackColor = System.Drawing.Color.Transparent;
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 75F));
            this.tableLayoutPanel1.Controls.Add(this.panel3, 1, 2);
            this.tableLayoutPanel1.Controls.Add(this.panel2, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.smS_Label1, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.smS_Label2, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.smS_Label3, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.txtTANCD, 1, 1);
            this.tableLayoutPanel1.Location = new System.Drawing.Point(12, 84);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 3;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(440, 123);
            this.tableLayoutPanel1.TabIndex = 1;
            // 
            // panel3
            // 
            this.panel3.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.panel3.Controls.Add(this.txtPASWD);
            this.panel3.Controls.Add(this.btnChange);
            this.panel3.Location = new System.Drawing.Point(110, 91);
            this.panel3.Margin = new System.Windows.Forms.Padding(0);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(283, 23);
            this.panel3.TabIndex = 3;
            this.panel3.TabStop = true;
            // 
            // txtPASWD
            // 
            this.txtPASWD.Decimalplace = ((byte)(0));
            this.txtPASWD.Location = new System.Drawing.Point(3, 3);
            this.txtPASWD.MaxLength = 2;
            this.txtPASWD.Name = "txtPASWD";
            this.txtPASWD.Size = new System.Drawing.Size(150, 19);
            this.txtPASWD.TabIndex = 1;
            this.txtPASWD.txt_Type = SMS.CustomControls.SMS_TextBox.Type.Normal;
            this.txtPASWD.UseMinus = false;
            this.txtPASWD.UseThousandSeparator = true;
            // 
            // btnChange
            // 
            this.btnChange.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.btnChange.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnChange.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnChange.Location = new System.Drawing.Point(154, 3);
            this.btnChange.Name = "btnChange";
            this.btnChange.Size = new System.Drawing.Size(119, 19);
            this.btnChange.TabIndex = 4;
            this.btnChange.TabStop = false;
            this.btnChange.Text = "パスワード変更";
            this.btnChange.UseVisualStyleBackColor = false;
            this.btnChange.Visible = false;
            // 
            // smS_Label1
            // 
            this.smS_Label1.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.smS_Label1.AutoSize = true;
            this.smS_Label1.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label1.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.smS_Label1.Location = new System.Drawing.Point(62, 14);
            this.smS_Label1.Name = "smS_Label1";
            this.smS_Label1.Size = new System.Drawing.Size(45, 12);
            this.smS_Label1.TabIndex = 0;
            this.smS_Label1.Text = "会社CD";
            // 
            // smS_Label2
            // 
            this.smS_Label2.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.smS_Label2.AutoSize = true;
            this.smS_Label2.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label2.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.smS_Label2.Location = new System.Drawing.Point(23, 55);
            this.smS_Label2.Name = "smS_Label2";
            this.smS_Label2.Size = new System.Drawing.Size(84, 12);
            this.smS_Label2.TabIndex = 1;
            this.smS_Label2.Text = "オペレータCD";
            // 
            // smS_Label3
            // 
            this.smS_Label3.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.smS_Label3.AutoSize = true;
            this.smS_Label3.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label3.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.smS_Label3.Location = new System.Drawing.Point(37, 96);
            this.smS_Label3.Name = "smS_Label3";
            this.smS_Label3.Size = new System.Drawing.Size(70, 12);
            this.smS_Label3.TabIndex = 3;
            this.smS_Label3.Text = "パスワード";
            // 
            // txtTANCD
            // 
            this.txtTANCD.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.txtTANCD.Decimalplace = ((byte)(0));
            this.txtTANCD.Location = new System.Drawing.Point(113, 52);
            this.txtTANCD.MaxLength = 2;
            this.txtTANCD.Name = "txtTANCD";
            this.txtTANCD.Size = new System.Drawing.Size(44, 19);
            this.txtTANCD.TabIndex = 2;
            this.txtTANCD.txt_Type = SMS.CustomControls.SMS_TextBox.Type.IntegerOnly;
            this.txtTANCD.UseMinus = false;
            this.txtTANCD.UseThousandSeparator = true;
            // 
            // panel4
            // 
            this.panel4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel4.Controls.Add(this.panel5);
            this.panel4.Controls.Add(this.tableLayoutPanel2);
            this.panel4.Controls.Add(this.tableLayoutPanel1);
            this.panel4.Controls.Add(this.btnF12);
            this.panel4.Controls.Add(this.btnF1);
            this.panel4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel4.Location = new System.Drawing.Point(0, 0);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(465, 304);
            this.panel4.TabIndex = 2;
            this.panel4.TabStop = true;
            // 
            // panel5
            // 
            this.panel5.Controls.Add(this.bnF1);
            this.panel5.Controls.Add(this.bnF12);
            this.panel5.Location = new System.Drawing.Point(12, 244);
            this.panel5.Name = "panel5";
            this.panel5.Size = new System.Drawing.Size(438, 41);
            this.panel5.TabIndex = 5;
            // 
            // bnF1
            // 
            this.bnF1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.bnF1.Cursor = System.Windows.Forms.Cursors.Hand;
            this.bnF1.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.bnF1.Location = new System.Drawing.Point(11, 6);
            this.bnF1.Name = "bnF1";
            this.bnF1.Size = new System.Drawing.Size(104, 28);
            this.bnF1.TabIndex = 1;
            this.bnF1.Text = "終了(F1)";
            this.bnF1.UseVisualStyleBackColor = false;
            this.bnF1.Click += new System.EventHandler(this.bnF1_Click_2);
            // 
            // bnF12
            // 
            this.bnF12.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.bnF12.Cursor = System.Windows.Forms.Cursors.Hand;
            this.bnF12.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.bnF12.Location = new System.Drawing.Point(317, 6);
            this.bnF12.Name = "bnF12";
            this.bnF12.Size = new System.Drawing.Size(112, 28);
            this.bnF12.TabIndex = 0;
            this.bnF12.Text = "ログイン(F12)";
            this.bnF12.UseVisualStyleBackColor = false;
            this.bnF12.Click += new System.EventHandler(this.bnF12_Click_2);
            // 
            // tableLayoutPanel2
            // 
            this.tableLayoutPanel2.ColumnCount = 2;
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel2.Location = new System.Drawing.Point(14, 255);
            this.tableLayoutPanel2.Name = "tableLayoutPanel2";
            this.tableLayoutPanel2.RowCount = 1;
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 29F));
            this.tableLayoutPanel2.Size = new System.Drawing.Size(436, 29);
            this.tableLayoutPanel2.TabIndex = 4;
            // 
            // btnF12
            // 
            this.btnF12.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.btnF12.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnF12.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnF12.Location = new System.Drawing.Point(354, 5);
            this.btnF12.Name = "btnF12";
            this.btnF12.Size = new System.Drawing.Size(98, 30);
            this.btnF12.TabIndex = 0;
            this.btnF12.Text = "ログイン(F12)";
            this.btnF12.UseVisualStyleBackColor = false;
            // 
            // btnF1
            // 
            this.btnF1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.btnF1.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnF1.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnF1.Location = new System.Drawing.Point(11, 5);
            this.btnF1.Name = "btnF1";
            this.btnF1.Size = new System.Drawing.Size(98, 30);
            this.btnF1.TabIndex = 1;
            this.btnF1.Text = "終了(F1)";
            this.btnF1.UseVisualStyleBackColor = false;
            // 
            // frmLogin
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            this.ClientSize = new System.Drawing.Size(465, 304);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.panel4);
            this.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmLogin";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Login";
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmLogin_KeyDown);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.panel4.ResumeLayout(false);
            this.panel5.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lblSProcessingDate;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Panel panel4;
        private CustomControls.SMS_TextBox txtKAICD;
        private CustomControls.SMS_TextBox txtPASWD;
        private CustomControls.SMS_Button btnChange;
        private CustomControls.SMS_Label smS_Label1;
        private CustomControls.SMS_Label smS_Label2;
        private CustomControls.SMS_Label smS_Label3;
        private CustomControls.SMS_TextBox txtTANCD;
        private CustomControls.SMS_Button btnF12;
        private CustomControls.SMS_Button btnF1;
        private System.Windows.Forms.Label lblJSYRN;
        private CustomControls.SMS_Button bnF12;
        private CustomControls.SMS_Button bnF1;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel2;
        private System.Windows.Forms.Panel panel5;
        private System.Windows.Forms.Label lblVersion;
    }
}