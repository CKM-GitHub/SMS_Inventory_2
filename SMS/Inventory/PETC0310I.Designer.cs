namespace SMS.Inventory
{
    partial class frm0310I
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
            this.txtInput = new SMS.CustomControls.SMS_TextBox();
            this.smS_Label2 = new SMS.CustomControls.SMS_Label();
            this.button1 = new System.Windows.Forms.Button();
            this.txtProcessNo = new SMS.CustomControls.SMS_TextBox();
            this.txtSlipNo = new SMS.CustomControls.SMS_TextBox();
            this.smS_Label3 = new SMS.CustomControls.SMS_Label();
            this.smS_Label4 = new SMS.CustomControls.SMS_Label();
            this.smS_Label5 = new SMS.CustomControls.SMS_Label();
            this.smS_Button1 = new SMS.CustomControls.SMS_Button();
            this.txtDate = new CKM_Controls.CKM_TextBox();
            this.panelBody.SuspendLayout();
            this.panelTitleLeft.SuspendLayout();
            this.panelTitle.SuspendLayout();
            this.SuspendLayout();
            // 
            // panelBody
            // 
            this.panelBody.Controls.Add(this.txtDate);
            this.panelBody.Controls.Add(this.smS_Button1);
            this.panelBody.Controls.Add(this.smS_Label5);
            this.panelBody.Controls.Add(this.smS_Label4);
            this.panelBody.Controls.Add(this.smS_Label3);
            this.panelBody.Controls.Add(this.txtSlipNo);
            this.panelBody.Controls.Add(this.txtProcessNo);
            this.panelBody.Controls.Add(this.button1);
            this.panelBody.Controls.Add(this.smS_Label2);
            this.panelBody.Controls.Add(this.txtInput);
            // 
            // txtInput
            // 
            this.txtInput.Decimalplace = ((byte)(0));
            this.txtInput.Enabled = false;
            this.txtInput.Location = new System.Drawing.Point(157, 42);
            this.txtInput.Name = "txtInput";
            this.txtInput.Size = new System.Drawing.Size(400, 19);
            this.txtInput.TabIndex = 0;
            this.txtInput.txt_Type = SMS.CustomControls.SMS_TextBox.Type.Normal;
            this.txtInput.UseMinus = false;
            this.txtInput.UseThousandSeparator = true;
            // 
            // smS_Label2
            // 
            this.smS_Label2.AutoSize = true;
            this.smS_Label2.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label2.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label2.Location = new System.Drawing.Point(60, 45);
            this.smS_Label2.Name = "smS_Label2";
            this.smS_Label2.Size = new System.Drawing.Size(83, 12);
            this.smS_Label2.TabIndex = 1;
            this.smS_Label2.Text = "取込フォルダ";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(556, 41);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(25, 22);
            this.button1.TabIndex = 0;
            this.button1.Text = "▼";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // txtProcessNo
            // 
            this.txtProcessNo.Decimalplace = ((byte)(0));
            this.txtProcessNo.Enabled = false;
            this.txtProcessNo.Location = new System.Drawing.Point(223, 77);
            this.txtProcessNo.MaxLength = 6;
            this.txtProcessNo.Name = "txtProcessNo";
            this.txtProcessNo.Size = new System.Drawing.Size(100, 19);
            this.txtProcessNo.TabIndex = 1;
            this.txtProcessNo.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtProcessNo.txt_Type = SMS.CustomControls.SMS_TextBox.Type.IntegerOnly;
            this.txtProcessNo.UseMinus = false;
            this.txtProcessNo.UseThousandSeparator = false;
            // 
            // txtSlipNo
            // 
            this.txtSlipNo.Decimalplace = ((byte)(0));
            this.txtSlipNo.Enabled = false;
            this.txtSlipNo.Location = new System.Drawing.Point(223, 127);
            this.txtSlipNo.MaxLength = 3;
            this.txtSlipNo.Name = "txtSlipNo";
            this.txtSlipNo.Size = new System.Drawing.Size(100, 19);
            this.txtSlipNo.TabIndex = 3;
            this.txtSlipNo.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtSlipNo.txt_Type = SMS.CustomControls.SMS_TextBox.Type.Normal;
            this.txtSlipNo.UseMinus = false;
            this.txtSlipNo.UseThousandSeparator = true;
            // 
            // smS_Label3
            // 
            this.smS_Label3.AutoSize = true;
            this.smS_Label3.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label3.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label3.Location = new System.Drawing.Point(156, 80);
            this.smS_Label3.Name = "smS_Label3";
            this.smS_Label3.Size = new System.Drawing.Size(57, 12);
            this.smS_Label3.TabIndex = 6;
            this.smS_Label3.Text = "処理番号";
            // 
            // smS_Label4
            // 
            this.smS_Label4.AutoSize = true;
            this.smS_Label4.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label4.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label4.Location = new System.Drawing.Point(156, 105);
            this.smS_Label4.Name = "smS_Label4";
            this.smS_Label4.Size = new System.Drawing.Size(57, 12);
            this.smS_Label4.TabIndex = 7;
            this.smS_Label4.Text = "伝票日付";
            // 
            // smS_Label5
            // 
            this.smS_Label5.AutoSize = true;
            this.smS_Label5.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label5.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label5.Location = new System.Drawing.Point(156, 130);
            this.smS_Label5.Name = "smS_Label5";
            this.smS_Label5.Size = new System.Drawing.Size(57, 12);
            this.smS_Label5.TabIndex = 8;
            this.smS_Label5.Text = "伝票番号";
            // 
            // smS_Button1
            // 
            this.smS_Button1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.smS_Button1.Cursor = System.Windows.Forms.Cursors.Hand;
            this.smS_Button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.smS_Button1.Location = new System.Drawing.Point(1594, 45);
            this.smS_Button1.Name = "smS_Button1";
            this.smS_Button1.Size = new System.Drawing.Size(150, 25);
            this.smS_Button1.TabIndex = 4;
            this.smS_Button1.Text = "出力(F12)";
            this.smS_Button1.UseVisualStyleBackColor = false;
            this.smS_Button1.Click += new System.EventHandler(this.smS_Button1_Click);
            // 
            // txtDate
            // 
            this.txtDate.CKM_Reqired = true;
            this.txtDate.CKM_Type = CKM_Controls.CKM_TextBox.Type.Date;
            this.txtDate.DecimalPlace = 0;
            this.txtDate.Enabled = false;
            this.txtDate.Location = new System.Drawing.Point(223, 102);
            this.txtDate.Name = "txtDate";
            this.txtDate.Size = new System.Drawing.Size(100, 19);
            this.txtDate.TabIndex = 2;
            this.txtDate.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtDate.UseMinus = false;
            this.txtDate.UseThousandSeparator = false;
            // 
            // frm0310I
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1884, 951);
            this.KeyPreview = true;
            this.Location = new System.Drawing.Point(0, 0);
            this.Name = "frm0310I";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "PETC0310I";
            this.Load += new System.EventHandler(this.PETC0310I_Load);
            this.panelBody.ResumeLayout(false);
            this.panelBody.PerformLayout();
            this.panelTitleLeft.ResumeLayout(false);
            this.panelTitle.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private CustomControls.SMS_Label smS_Label5;
        private CustomControls.SMS_Label smS_Label4;
        private CustomControls.SMS_Label smS_Label3;
        private CustomControls.SMS_TextBox txtSlipNo;
        private CustomControls.SMS_TextBox txtProcessNo;
        private System.Windows.Forms.Button button1;
        private CustomControls.SMS_Label smS_Label2;
        private CustomControls.SMS_TextBox txtInput;
        private CustomControls.SMS_Button smS_Button1;
        private CKM_Controls.CKM_TextBox txtDate;
    }
}