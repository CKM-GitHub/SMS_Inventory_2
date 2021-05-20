namespace SMS.Inventory
{
    partial class frmPETC0307I
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
            this.smsLabel3 = new SMS.CustomControls.SMS_Label();
            this.smS_Label4 = new SMS.CustomControls.SMS_Label();
            this.cboRakutenStore = new SMS.CustomControls.SMS_ComboBox();
            this.txtStartDate = new CKM_Controls.CKM_TextBox();
            this.txtEndDate = new CKM_Controls.CKM_TextBox();
            this.panelBody.SuspendLayout();
            this.panelTitleLeft.SuspendLayout();
            this.panelTitle.SuspendLayout();
            this.SuspendLayout();
            // 
            // panelBody
            // 
            this.panelBody.Controls.Add(this.txtEndDate);
            this.panelBody.Controls.Add(this.txtStartDate);
            this.panelBody.Controls.Add(this.cboRakutenStore);
            this.panelBody.Controls.Add(this.smS_Label4);
            this.panelBody.Controls.Add(this.smsLabel3);
            this.panelBody.Controls.Add(this.smS_Label2);
            // 
            // smS_Label2
            // 
            this.smS_Label2.AutoSize = true;
            this.smS_Label2.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label2.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label2.Location = new System.Drawing.Point(52, 18);
            this.smS_Label2.Name = "smS_Label2";
            this.smS_Label2.Size = new System.Drawing.Size(97, 12);
            this.smS_Label2.TabIndex = 0;
            this.smS_Label2.Text = "(すご楽)出荷日";
            // 
            // smsLabel3
            // 
            this.smsLabel3.AutoSize = true;
            this.smsLabel3.BackColor = System.Drawing.Color.Transparent;
            this.smsLabel3.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smsLabel3.Location = new System.Drawing.Point(254, 18);
            this.smsLabel3.Name = "smsLabel3";
            this.smsLabel3.Size = new System.Drawing.Size(31, 12);
            this.smsLabel3.TabIndex = 1;
            this.smsLabel3.Text = "から";
            // 
            // smS_Label4
            // 
            this.smS_Label4.AutoSize = true;
            this.smS_Label4.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label4.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label4.Location = new System.Drawing.Point(104, 53);
            this.smS_Label4.Name = "smS_Label4";
            this.smS_Label4.Size = new System.Drawing.Size(44, 12);
            this.smS_Label4.TabIndex = 2;
            this.smS_Label4.Text = "楽天店";
            // 
            // cboRakutenStore
            // 
            this.cboRakutenStore.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource;
            this.cboRakutenStore.Cbo_Type = SMS.CustomControls.SMS_ComboBox.cbo_Type.m_nendo;
            this.cboRakutenStore.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cboRakutenStore.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboRakutenStore.FormattingEnabled = true;
            this.cboRakutenStore.Location = new System.Drawing.Point(151, 49);
            this.cboRakutenStore.Name = "cboRakutenStore";
            this.cboRakutenStore.Size = new System.Drawing.Size(200, 20);
            this.cboRakutenStore.TabIndex = 6;
            // 
            // txtStartDate
            // 
            this.txtStartDate.CKM_Reqired = true;
            this.txtStartDate.CKM_Type = CKM_Controls.CKM_TextBox.Type.Date;
            this.txtStartDate.DecimalPlace = 0;
            this.txtStartDate.Location = new System.Drawing.Point(151, 15);
            this.txtStartDate.Name = "txtStartDate";
            this.txtStartDate.Size = new System.Drawing.Size(100, 19);
            this.txtStartDate.TabIndex = 7;
            this.txtStartDate.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtStartDate.UseMinus = true;
            this.txtStartDate.UseThousandSeparator = true;
            // 
            // txtEndDate
            // 
            this.txtEndDate.CKM_Reqired = true;
            this.txtEndDate.CKM_Type = CKM_Controls.CKM_TextBox.Type.Date;
            this.txtEndDate.DecimalPlace = 0;
            this.txtEndDate.Location = new System.Drawing.Point(288, 15);
            this.txtEndDate.Name = "txtEndDate";
            this.txtEndDate.Size = new System.Drawing.Size(100, 19);
            this.txtEndDate.TabIndex = 8;
            this.txtEndDate.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtEndDate.UseMinus = true;
            this.txtEndDate.UseThousandSeparator = true;
            // 
            // frmPETC0307I
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1884, 951);
            this.KeyPreview = true;
            this.Location = new System.Drawing.Point(0, 0);
            this.Name = "frmPETC0307I";
            this.Text = "PECT0307I";
            this.Load += new System.EventHandler(this.frmPECT0307I_Load);
            this.panelBody.ResumeLayout(false);
            this.panelBody.PerformLayout();
            this.panelTitleLeft.ResumeLayout(false);
            this.panelTitle.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private CustomControls.SMS_Label smS_Label2;
        private CustomControls.SMS_Label smS_Label4;
        private CustomControls.SMS_Label smsLabel3;
        private CustomControls.SMS_ComboBox cboRakutenStore;
        private CKM_Controls.CKM_TextBox txtEndDate;
        private CKM_Controls.CKM_TextBox txtStartDate;

    }
}