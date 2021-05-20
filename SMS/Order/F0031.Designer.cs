namespace SMS.Order
{
    partial class F0031
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
            this.dateTextBox1 = new SMS.CustomControls.DateTextBox();
            this.dateTextBox2 = new SMS.CustomControls.DateTextBox();
            this.checkBox2 = new System.Windows.Forms.CheckBox();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.smS_Label5 = new SMS.CustomControls.SMS_Label();
            this.smS_Label4 = new SMS.CustomControls.SMS_Label();
            this.smS_Label3 = new SMS.CustomControls.SMS_Label();
            this.smS_Label2 = new SMS.CustomControls.SMS_Label();
            this.smS_Label6 = new SMS.CustomControls.SMS_Label();
            this.combostore = new System.Windows.Forms.ComboBox();
            this.uC_Search1 = new SMS.CustomControls.UC_Search();
            this.panelBody.SuspendLayout();
            this.panelTitleLeft.SuspendLayout();
            this.panelTitle.SuspendLayout();
            this.SuspendLayout();
            // 
            // panelBody
            // 
            this.panelBody.Controls.Add(this.uC_Search1);
            this.panelBody.Controls.Add(this.smS_Label6);
            this.panelBody.Controls.Add(this.combostore);
            this.panelBody.Controls.Add(this.dateTextBox1);
            this.panelBody.Controls.Add(this.dateTextBox2);
            this.panelBody.Controls.Add(this.checkBox2);
            this.panelBody.Controls.Add(this.checkBox1);
            this.panelBody.Controls.Add(this.smS_Label5);
            this.panelBody.Controls.Add(this.smS_Label4);
            this.panelBody.Controls.Add(this.smS_Label3);
            this.panelBody.Controls.Add(this.smS_Label2);
            this.panelBody.TabIndex = 0;
            // 
            // dateTextBox1
            // 
            this.dateTextBox1.Location = new System.Drawing.Point(266, 31);
            this.dateTextBox1.Mask = "0000/00/00";
            this.dateTextBox1.MaskFormat = "yyyy/MM/dd";
            this.dateTextBox1.Name = "dateTextBox1";
            this.dateTextBox1.PromptChar = ' ';
            this.dateTextBox1.Size = new System.Drawing.Size(100, 19);
            this.dateTextBox1.TabIndex = 1;
            // 
            // dateTextBox2
            // 
            this.dateTextBox2.Location = new System.Drawing.Point(136, 31);
            this.dateTextBox2.Mask = "0000/00/00";
            this.dateTextBox2.MaskFormat = "yyyy/MM/dd";
            this.dateTextBox2.Name = "dateTextBox2";
            this.dateTextBox2.PromptChar = ' ';
            this.dateTextBox2.Size = new System.Drawing.Size(100, 19);
            this.dateTextBox2.TabIndex = 0;
            // 
            // checkBox2
            // 
            this.checkBox2.AutoSize = true;
            this.checkBox2.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBox2.Location = new System.Drawing.Point(228, 95);
            this.checkBox2.Name = "checkBox2";
            this.checkBox2.Size = new System.Drawing.Size(63, 16);
            this.checkBox2.TabIndex = 5;
            this.checkBox2.Text = "出力済";
            this.checkBox2.UseVisualStyleBackColor = true;
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBox1.Location = new System.Drawing.Point(136, 93);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(63, 16);
            this.checkBox1.TabIndex = 4;
            this.checkBox1.Text = "未出力";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // smS_Label5
            // 
            this.smS_Label5.AutoSize = true;
            this.smS_Label5.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label5.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label5.Location = new System.Drawing.Point(242, 34);
            this.smS_Label5.Name = "smS_Label5";
            this.smS_Label5.Size = new System.Drawing.Size(18, 12);
            this.smS_Label5.TabIndex = 186;
            this.smS_Label5.Text = "～";
            // 
            // smS_Label4
            // 
            this.smS_Label4.AutoSize = true;
            this.smS_Label4.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label4.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label4.Location = new System.Drawing.Point(76, 95);
            this.smS_Label4.Name = "smS_Label4";
            this.smS_Label4.Size = new System.Drawing.Size(57, 12);
            this.smS_Label4.TabIndex = 4;
            this.smS_Label4.Text = "出力対象";
            // 
            // smS_Label3
            // 
            this.smS_Label3.AutoSize = true;
            this.smS_Label3.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label3.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label3.Location = new System.Drawing.Point(90, 67);
            this.smS_Label3.Name = "smS_Label3";
            this.smS_Label3.Size = new System.Drawing.Size(44, 12);
            this.smS_Label3.TabIndex = 184;
            this.smS_Label3.Text = "仕入先";
            // 
            // smS_Label2
            // 
            this.smS_Label2.AutoSize = true;
            this.smS_Label2.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label2.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label2.Location = new System.Drawing.Point(89, 34);
            this.smS_Label2.Name = "smS_Label2";
            this.smS_Label2.Size = new System.Drawing.Size(44, 12);
            this.smS_Label2.TabIndex = 183;
            this.smS_Label2.Text = "仕入日";
            // 
            // smS_Label6
            // 
            this.smS_Label6.AutoSize = true;
            this.smS_Label6.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label6.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label6.Location = new System.Drawing.Point(1720, 26);
            this.smS_Label6.Name = "smS_Label6";
            this.smS_Label6.Size = new System.Drawing.Size(31, 12);
            this.smS_Label6.TabIndex = 192;
            this.smS_Label6.Text = "店舗";
            // 
            // combostore
            // 
            this.combostore.FormattingEnabled = true;
            this.combostore.Location = new System.Drawing.Point(1757, 23);
            this.combostore.Name = "combostore";
            this.combostore.Size = new System.Drawing.Size(115, 20);
            this.combostore.TabIndex = 2;
            // 
            // uC_Search1
            // 
            this.uC_Search1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            this.uC_Search1.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.uC_Search1.Location = new System.Drawing.Point(136, 61);
            this.uC_Search1.Margin = new System.Windows.Forms.Padding(0);
            this.uC_Search1.Name = "uC_Search1";
            this.uC_Search1.Size = new System.Drawing.Size(460, 23);
            this.uC_Search1.TabIndex = 3;
            this.uC_Search1.UC_Code = "";
            this.uC_Search1.UC_Code_Width = 72;
            this.uC_Search1.UC_Flag = 0;
            this.uC_Search1.UC_Name = "";
            this.uC_Search1.UC_Name_Width = 350;
            this.uC_Search1.UC_NameVisible = true;
            this.uC_Search1.UC_SearchEnable = true;
            this.uC_Search1.UC_Type = SMS.CustomControls.UC_Search.Type.Maker;
            // 
            // F0031
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1884, 951);
            this.F12_Text = "出力(F12)";
            this.KeyPreview = true;
            this.Location = new System.Drawing.Point(0, 0);
            this.Name = "F0031";
            this.Text = "F0031";
            this.Load += new System.EventHandler(this.F0031_Load);
            this.panelBody.ResumeLayout(false);
            this.panelBody.PerformLayout();
            this.panelTitleLeft.ResumeLayout(false);
            this.panelTitle.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private CustomControls.DateTextBox dateTextBox1;
        private CustomControls.DateTextBox dateTextBox2;
        private System.Windows.Forms.CheckBox checkBox2;
        private System.Windows.Forms.CheckBox checkBox1;
        private CustomControls.SMS_Label smS_Label5;
        private CustomControls.SMS_Label smS_Label4;
        private CustomControls.SMS_Label smS_Label3;
        private CustomControls.SMS_Label smS_Label2;
        private CustomControls.SMS_Label smS_Label6;
        private System.Windows.Forms.ComboBox combostore;
        private CustomControls.UC_Search uC_Search1;
    }
}