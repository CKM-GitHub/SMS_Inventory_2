namespace SMS.Inventory
{
    partial class frmPETC0301I
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
            this.smS_Label3 = new SMS.CustomControls.SMS_Label();
            this.cbo_YahooShop = new System.Windows.Forms.ComboBox();
            this.txtStartDate = new CKM_Controls.CKM_TextBox();
            this.smS_Label4 = new SMS.CustomControls.SMS_Label();
            this.txtEndDate = new CKM_Controls.CKM_TextBox();
            this.panelBody.SuspendLayout();
            this.panelTitleLeft.SuspendLayout();
            this.panelTitle.SuspendLayout();
            this.SuspendLayout();
            // 
            // panelBody
            // 
            this.panelBody.Controls.Add(this.txtEndDate);
            this.panelBody.Controls.Add(this.smS_Label4);
            this.panelBody.Controls.Add(this.txtStartDate);
            this.panelBody.Controls.Add(this.cbo_YahooShop);
            this.panelBody.Controls.Add(this.smS_Label3);
            this.panelBody.Controls.Add(this.smS_Label2);
            // 
            // smS_Label2
            // 
            this.smS_Label2.AutoSize = true;
            this.smS_Label2.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label2.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label2.Location = new System.Drawing.Point(40, 22);
            this.smS_Label2.Name = "smS_Label2";
            this.smS_Label2.Size = new System.Drawing.Size(97, 12);
            this.smS_Label2.TabIndex = 0;
            this.smS_Label2.Text = "(すご楽)請求日";
            // 
            // smS_Label3
            // 
            this.smS_Label3.AutoSize = true;
            this.smS_Label3.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label3.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label3.Location = new System.Drawing.Point(82, 50);
            this.smS_Label3.Name = "smS_Label3";
            this.smS_Label3.Size = new System.Drawing.Size(53, 12);
            this.smS_Label3.TabIndex = 1;
            this.smS_Label3.Text = "Yahoo店";
            // 
            // cbo_YahooShop
            // 
            this.cbo_YahooShop.FormattingEnabled = true;
            this.cbo_YahooShop.Location = new System.Drawing.Point(142, 47);
            this.cbo_YahooShop.Name = "cbo_YahooShop";
            this.cbo_YahooShop.Size = new System.Drawing.Size(200, 20);
            this.cbo_YahooShop.TabIndex = 3;
            // 
            // txtStartDate
            // 
            this.txtStartDate.CKM_Reqired = true;
            this.txtStartDate.CKM_Type = CKM_Controls.CKM_TextBox.Type.Date;
            this.txtStartDate.Location = new System.Drawing.Point(143, 19);
            this.txtStartDate.Name = "txtStartDate";
            this.txtStartDate.Size = new System.Drawing.Size(100, 19);
            this.txtStartDate.TabIndex = 4;
            this.txtStartDate.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtStartDate.UseThousandSeparator = true;
            // 
            // smS_Label4
            // 
            this.smS_Label4.AutoSize = true;
            this.smS_Label4.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label4.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label4.Location = new System.Drawing.Point(249, 22);
            this.smS_Label4.Name = "smS_Label4";
            this.smS_Label4.Size = new System.Drawing.Size(31, 12);
            this.smS_Label4.TabIndex = 6;
            this.smS_Label4.Text = "から";
            // 
            // txtEndDate
            // 
            this.txtEndDate.CKM_Reqired = true;
            this.txtEndDate.CKM_Type = CKM_Controls.CKM_TextBox.Type.Date;
            this.txtEndDate.Location = new System.Drawing.Point(283, 19);
            this.txtEndDate.Name = "txtEndDate";
            this.txtEndDate.Size = new System.Drawing.Size(100, 19);
            this.txtEndDate.TabIndex = 7;
            this.txtEndDate.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtEndDate.UseThousandSeparator = true;
            // 
            // frmPETC0301I
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1884, 951);
            this.KeyPreview = true;
            this.Location = new System.Drawing.Point(0, 0);
            this.Name = "frmPETC0301I";
            this.Text = "PETC0301I";
            this.Load += new System.EventHandler(this.frmPETC0301I_Load);
            this.panelBody.ResumeLayout(false);
            this.panelBody.PerformLayout();
            this.panelTitleLeft.ResumeLayout(false);
            this.panelTitle.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private CustomControls.SMS_Label smS_Label3;
        private CustomControls.SMS_Label smS_Label2;
        private System.Windows.Forms.ComboBox cbo_YahooShop;
        private CKM_Controls.CKM_TextBox txtStartDate;
        private CustomControls.SMS_Label smS_Label4;
        private CKM_Controls.CKM_TextBox txtEndDate;
    }
}