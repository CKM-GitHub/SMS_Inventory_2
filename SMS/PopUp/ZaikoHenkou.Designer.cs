namespace SMS.PopUp
{
    partial class frmZaikoHenkou
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmZaikoHenkou));
            this.lblFormName = new System.Windows.Forms.Label();
            this.btnF12 = new SMS.CustomControls.SMS_Button();
            this.btnF1 = new SMS.CustomControls.SMS_Button();
            this.txtSKUCD = new SMS.CustomControls.SMS_TextBox();
            this.txtJANCD = new SMS.CustomControls.SMS_TextBox();
            this.txtItemName = new SMS.CustomControls.SMS_TextBox();
            this.txtColor = new SMS.CustomControls.SMS_TextBox();
            this.txtSize = new SMS.CustomControls.SMS_TextBox();
            this.txtCapitalZaiko = new SMS.CustomControls.SMS_TextBox();
            this.txtToyonakaZaiko = new SMS.CustomControls.SMS_TextBox();
            this.txtEsakaZaiko = new SMS.CustomControls.SMS_TextBox();
            this.txtSannomiyaZaiko = new SMS.CustomControls.SMS_TextBox();
            this.smS_Label10 = new SMS.CustomControls.SMS_Label();
            this.smS_Label9 = new SMS.CustomControls.SMS_Label();
            this.smS_Label8 = new SMS.CustomControls.SMS_Label();
            this.smS_Label7 = new SMS.CustomControls.SMS_Label();
            this.smS_Label6 = new SMS.CustomControls.SMS_Label();
            this.smS_Label5 = new SMS.CustomControls.SMS_Label();
            this.smS_Label4 = new SMS.CustomControls.SMS_Label();
            this.smS_Label3 = new SMS.CustomControls.SMS_Label();
            this.smS_Label2 = new SMS.CustomControls.SMS_Label();
            this.smS_Label1 = new SMS.CustomControls.SMS_Label();
            this.txtIshibashiZaiko = new SMS.CustomControls.SMS_TextBox();
            this.SuspendLayout();
            // 
            // lblFormName
            // 
            this.lblFormName.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(173)))), ((int)(((byte)(71)))));
            this.lblFormName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblFormName.Font = new System.Drawing.Font("MS Gothic", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFormName.Location = new System.Drawing.Point(89, 9);
            this.lblFormName.Name = "lblFormName";
            this.lblFormName.Size = new System.Drawing.Size(343, 22);
            this.lblFormName.TabIndex = 3;
            this.lblFormName.Text = "在庫変更";
            this.lblFormName.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btnF12
            // 
            this.btnF12.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.btnF12.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnF12.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnF12.Location = new System.Drawing.Point(414, 286);
            this.btnF12.Name = "btnF12";
            this.btnF12.Size = new System.Drawing.Size(100, 28);
            this.btnF12.TabIndex = 11;
            this.btnF12.Text = "確定(F12)";
            this.btnF12.UseVisualStyleBackColor = false;
            this.btnF12.Click += new System.EventHandler(this.btnF12_Click);
            // 
            // btnF1
            // 
            this.btnF1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.btnF1.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnF1.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnF1.Location = new System.Drawing.Point(11, 286);
            this.btnF1.Name = "btnF1";
            this.btnF1.Size = new System.Drawing.Size(100, 28);
            this.btnF1.TabIndex = 10;
            this.btnF1.Text = "戻る(F1)";
            this.btnF1.UseVisualStyleBackColor = false;
            this.btnF1.Click += new System.EventHandler(this.btnF1_Click);
            // 
            // txtSKUCD
            // 
            this.txtSKUCD.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtSKUCD.Decimalplace = ((byte)(0));
            this.txtSKUCD.Enabled = false;
            this.txtSKUCD.Location = new System.Drawing.Point(102, 44);
            this.txtSKUCD.Name = "txtSKUCD";
            this.txtSKUCD.Size = new System.Drawing.Size(175, 19);
            this.txtSKUCD.TabIndex = 0;
            this.txtSKUCD.txt_Type = SMS.CustomControls.SMS_TextBox.Type.Normal;
            this.txtSKUCD.UseMinus = false;
            this.txtSKUCD.UseThousandSeparator = true;
            // 
            // txtJANCD
            // 
            this.txtJANCD.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtJANCD.Decimalplace = ((byte)(0));
            this.txtJANCD.Enabled = false;
            this.txtJANCD.Location = new System.Drawing.Point(102, 67);
            this.txtJANCD.Name = "txtJANCD";
            this.txtJANCD.Size = new System.Drawing.Size(119, 19);
            this.txtJANCD.TabIndex = 1;
            this.txtJANCD.txt_Type = SMS.CustomControls.SMS_TextBox.Type.IntegerOnly;
            this.txtJANCD.UseMinus = false;
            this.txtJANCD.UseThousandSeparator = true;
            // 
            // txtItemName
            // 
            this.txtItemName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtItemName.Decimalplace = ((byte)(0));
            this.txtItemName.Enabled = false;
            this.txtItemName.Location = new System.Drawing.Point(102, 90);
            this.txtItemName.Name = "txtItemName";
            this.txtItemName.Size = new System.Drawing.Size(412, 19);
            this.txtItemName.TabIndex = 2;
            this.txtItemName.txt_Type = SMS.CustomControls.SMS_TextBox.Type.Normal;
            this.txtItemName.UseMinus = false;
            this.txtItemName.UseThousandSeparator = true;
            // 
            // txtColor
            // 
            this.txtColor.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtColor.Decimalplace = ((byte)(0));
            this.txtColor.Enabled = false;
            this.txtColor.Location = new System.Drawing.Point(102, 114);
            this.txtColor.Name = "txtColor";
            this.txtColor.Size = new System.Drawing.Size(202, 19);
            this.txtColor.TabIndex = 3;
            this.txtColor.txt_Type = SMS.CustomControls.SMS_TextBox.Type.Normal;
            this.txtColor.UseMinus = false;
            this.txtColor.UseThousandSeparator = true;
            // 
            // txtSize
            // 
            this.txtSize.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtSize.Decimalplace = ((byte)(0));
            this.txtSize.Enabled = false;
            this.txtSize.Location = new System.Drawing.Point(102, 137);
            this.txtSize.Name = "txtSize";
            this.txtSize.Size = new System.Drawing.Size(202, 19);
            this.txtSize.TabIndex = 4;
            this.txtSize.txt_Type = SMS.CustomControls.SMS_TextBox.Type.Normal;
            this.txtSize.UseMinus = false;
            this.txtSize.UseThousandSeparator = true;
            // 
            // txtCapitalZaiko
            // 
            this.txtCapitalZaiko.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtCapitalZaiko.Decimalplace = ((byte)(0));
            this.txtCapitalZaiko.Location = new System.Drawing.Point(102, 160);
            this.txtCapitalZaiko.Name = "txtCapitalZaiko";
            this.txtCapitalZaiko.Size = new System.Drawing.Size(55, 19);
            this.txtCapitalZaiko.TabIndex = 5;
            this.txtCapitalZaiko.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtCapitalZaiko.txt_Type = SMS.CustomControls.SMS_TextBox.Type.Decimal;
            this.txtCapitalZaiko.UseMinus = false;
            this.txtCapitalZaiko.UseThousandSeparator = true;
            // 
            // txtToyonakaZaiko
            // 
            this.txtToyonakaZaiko.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtToyonakaZaiko.Decimalplace = ((byte)(0));
            this.txtToyonakaZaiko.Location = new System.Drawing.Point(102, 183);
            this.txtToyonakaZaiko.Name = "txtToyonakaZaiko";
            this.txtToyonakaZaiko.Size = new System.Drawing.Size(55, 19);
            this.txtToyonakaZaiko.TabIndex = 6;
            this.txtToyonakaZaiko.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtToyonakaZaiko.txt_Type = SMS.CustomControls.SMS_TextBox.Type.Decimal;
            this.txtToyonakaZaiko.UseMinus = false;
            this.txtToyonakaZaiko.UseThousandSeparator = true;
            // 
            // txtEsakaZaiko
            // 
            this.txtEsakaZaiko.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtEsakaZaiko.Decimalplace = ((byte)(0));
            this.txtEsakaZaiko.Location = new System.Drawing.Point(102, 230);
            this.txtEsakaZaiko.Name = "txtEsakaZaiko";
            this.txtEsakaZaiko.Size = new System.Drawing.Size(55, 19);
            this.txtEsakaZaiko.TabIndex = 8;
            this.txtEsakaZaiko.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtEsakaZaiko.txt_Type = SMS.CustomControls.SMS_TextBox.Type.Decimal;
            this.txtEsakaZaiko.UseMinus = false;
            this.txtEsakaZaiko.UseThousandSeparator = true;
            // 
            // txtSannomiyaZaiko
            // 
            this.txtSannomiyaZaiko.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtSannomiyaZaiko.Decimalplace = ((byte)(0));
            this.txtSannomiyaZaiko.Location = new System.Drawing.Point(102, 254);
            this.txtSannomiyaZaiko.Name = "txtSannomiyaZaiko";
            this.txtSannomiyaZaiko.Size = new System.Drawing.Size(55, 19);
            this.txtSannomiyaZaiko.TabIndex = 9;
            this.txtSannomiyaZaiko.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtSannomiyaZaiko.txt_Type = SMS.CustomControls.SMS_TextBox.Type.Decimal;
            this.txtSannomiyaZaiko.UseMinus = false;
            this.txtSannomiyaZaiko.UseThousandSeparator = true;
            // 
            // smS_Label10
            // 
            this.smS_Label10.AutoSize = true;
            this.smS_Label10.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label10.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label10.Location = new System.Drawing.Point(30, 258);
            this.smS_Label10.Name = "smS_Label10";
            this.smS_Label10.Size = new System.Drawing.Size(70, 12);
            this.smS_Label10.TabIndex = 14;
            this.smS_Label10.Text = "三宮在庫数";
            // 
            // smS_Label9
            // 
            this.smS_Label9.AutoSize = true;
            this.smS_Label9.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label9.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label9.Location = new System.Drawing.Point(30, 234);
            this.smS_Label9.Name = "smS_Label9";
            this.smS_Label9.Size = new System.Drawing.Size(70, 12);
            this.smS_Label9.TabIndex = 13;
            this.smS_Label9.Text = "江坂在庫数";
            // 
            // smS_Label8
            // 
            this.smS_Label8.AutoSize = true;
            this.smS_Label8.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label8.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label8.Location = new System.Drawing.Point(30, 210);
            this.smS_Label8.Name = "smS_Label8";
            this.smS_Label8.Size = new System.Drawing.Size(70, 12);
            this.smS_Label8.TabIndex = 12;
            this.smS_Label8.Text = "石橋在庫数";
            // 
            // smS_Label7
            // 
            this.smS_Label7.AutoSize = true;
            this.smS_Label7.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label7.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label7.Location = new System.Drawing.Point(30, 187);
            this.smS_Label7.Name = "smS_Label7";
            this.smS_Label7.Size = new System.Drawing.Size(70, 12);
            this.smS_Label7.TabIndex = 11;
            this.smS_Label7.Text = "豊中在庫数";
            // 
            // smS_Label6
            // 
            this.smS_Label6.AutoSize = true;
            this.smS_Label6.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label6.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label6.Location = new System.Drawing.Point(30, 163);
            this.smS_Label6.Name = "smS_Label6";
            this.smS_Label6.Size = new System.Drawing.Size(70, 12);
            this.smS_Label6.TabIndex = 10;
            this.smS_Label6.Text = "自社在庫数";
            // 
            // smS_Label5
            // 
            this.smS_Label5.AutoSize = true;
            this.smS_Label5.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label5.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label5.Location = new System.Drawing.Point(56, 141);
            this.smS_Label5.Name = "smS_Label5";
            this.smS_Label5.Size = new System.Drawing.Size(44, 12);
            this.smS_Label5.TabIndex = 9;
            this.smS_Label5.Text = "サイズ";
            // 
            // smS_Label4
            // 
            this.smS_Label4.AutoSize = true;
            this.smS_Label4.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label4.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label4.Location = new System.Drawing.Point(56, 117);
            this.smS_Label4.Name = "smS_Label4";
            this.smS_Label4.Size = new System.Drawing.Size(44, 12);
            this.smS_Label4.TabIndex = 8;
            this.smS_Label4.Text = "カラー";
            // 
            // smS_Label3
            // 
            this.smS_Label3.AutoSize = true;
            this.smS_Label3.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label3.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label3.Location = new System.Drawing.Point(56, 94);
            this.smS_Label3.Name = "smS_Label3";
            this.smS_Label3.Size = new System.Drawing.Size(44, 12);
            this.smS_Label3.TabIndex = 7;
            this.smS_Label3.Text = "商品名";
            // 
            // smS_Label2
            // 
            this.smS_Label2.AutoSize = true;
            this.smS_Label2.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label2.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label2.Location = new System.Drawing.Point(60, 71);
            this.smS_Label2.Name = "smS_Label2";
            this.smS_Label2.Size = new System.Drawing.Size(40, 12);
            this.smS_Label2.TabIndex = 6;
            this.smS_Label2.Text = "JANCD";
            // 
            // smS_Label1
            // 
            this.smS_Label1.AutoSize = true;
            this.smS_Label1.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label1.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label1.Location = new System.Drawing.Point(60, 48);
            this.smS_Label1.Name = "smS_Label1";
            this.smS_Label1.Size = new System.Drawing.Size(40, 12);
            this.smS_Label1.TabIndex = 5;
            this.smS_Label1.Text = "SKUCD";
            // 
            // txtIshibashiZaiko
            // 
            this.txtIshibashiZaiko.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtIshibashiZaiko.Decimalplace = ((byte)(0));
            this.txtIshibashiZaiko.Location = new System.Drawing.Point(102, 206);
            this.txtIshibashiZaiko.Name = "txtIshibashiZaiko";
            this.txtIshibashiZaiko.Size = new System.Drawing.Size(55, 19);
            this.txtIshibashiZaiko.TabIndex = 7;
            this.txtIshibashiZaiko.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtIshibashiZaiko.txt_Type = SMS.CustomControls.SMS_TextBox.Type.Decimal;
            this.txtIshibashiZaiko.UseMinus = false;
            this.txtIshibashiZaiko.UseThousandSeparator = true;
            // 
            // frmZaikoHenkou
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            this.ClientSize = new System.Drawing.Size(534, 320);
            this.Controls.Add(this.btnF12);
            this.Controls.Add(this.btnF1);
            this.Controls.Add(this.txtSKUCD);
            this.Controls.Add(this.txtJANCD);
            this.Controls.Add(this.txtItemName);
            this.Controls.Add(this.txtColor);
            this.Controls.Add(this.txtSize);
            this.Controls.Add(this.txtCapitalZaiko);
            this.Controls.Add(this.txtToyonakaZaiko);
            this.Controls.Add(this.txtEsakaZaiko);
            this.Controls.Add(this.txtSannomiyaZaiko);
            this.Controls.Add(this.smS_Label10);
            this.Controls.Add(this.smS_Label9);
            this.Controls.Add(this.smS_Label8);
            this.Controls.Add(this.smS_Label7);
            this.Controls.Add(this.smS_Label6);
            this.Controls.Add(this.smS_Label5);
            this.Controls.Add(this.smS_Label4);
            this.Controls.Add(this.smS_Label3);
            this.Controls.Add(this.smS_Label2);
            this.Controls.Add(this.smS_Label1);
            this.Controls.Add(this.txtIshibashiZaiko);
            this.Controls.Add(this.lblFormName);
            this.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimizeBox = false;
            this.Name = "frmZaikoHenkou";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "ZaikoHenkou";
            this.Load += new System.EventHandler(this.frmZaikoHenkou_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmZaikoHenkou_KeyDown);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblFormName;
        private CustomControls.SMS_TextBox txtIshibashiZaiko;
        private CustomControls.SMS_Label smS_Label1;
        private CustomControls.SMS_Label smS_Label2;
        private CustomControls.SMS_Label smS_Label3;
        private CustomControls.SMS_Label smS_Label4;
        private CustomControls.SMS_Label smS_Label5;
        private CustomControls.SMS_Label smS_Label6;
        private CustomControls.SMS_Label smS_Label7;
        private CustomControls.SMS_Label smS_Label8;
        private CustomControls.SMS_Label smS_Label9;
        private CustomControls.SMS_Label smS_Label10;
        private CustomControls.SMS_TextBox txtSannomiyaZaiko;
        private CustomControls.SMS_TextBox txtEsakaZaiko;
        private CustomControls.SMS_TextBox txtToyonakaZaiko;
        private CustomControls.SMS_TextBox txtCapitalZaiko;
        private CustomControls.SMS_TextBox txtSize;
        private CustomControls.SMS_TextBox txtColor;
        private CustomControls.SMS_TextBox txtItemName;
        private CustomControls.SMS_TextBox txtJANCD;
        private CustomControls.SMS_TextBox txtSKUCD;
        private CustomControls.SMS_Button btnF12;
        private CustomControls.SMS_Button btnF1;
    }
}