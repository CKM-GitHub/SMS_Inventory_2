namespace SMS.Inventory
{
    partial class frmPSKS0114C_1
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
            this.btnEnd = new SMS.CustomControls.SMS_Button();
            this.btnStart = new SMS.CustomControls.SMS_Button();
            this.lblSKSMode = new System.Windows.Forms.Label();
            this.smS_Label6 = new SMS.CustomControls.SMS_Label();
            this.smS_Label5 = new SMS.CustomControls.SMS_Label();
            this.panelBody.SuspendLayout();
            this.panelTitleLeft.SuspendLayout();
            this.panelTitle.SuspendLayout();
            this.SuspendLayout();
            // 
            // panelBody
            // 
            this.panelBody.Controls.Add(this.smS_Label6);
            this.panelBody.Controls.Add(this.smS_Label5);
            this.panelBody.Controls.Add(this.btnEnd);
            this.panelBody.Controls.Add(this.btnStart);
            this.panelBody.Controls.Add(this.lblSKSMode);
            this.panelBody.TabIndex = 0;
            // 
            // btnEnd
            // 
            this.btnEnd.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.btnEnd.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnEnd.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnEnd.Location = new System.Drawing.Point(302, 42);
            this.btnEnd.Name = "btnEnd";
            this.btnEnd.Size = new System.Drawing.Size(109, 28);
            this.btnEnd.TabIndex = 1;
            this.btnEnd.Text = "停止";
            this.btnEnd.UseVisualStyleBackColor = false;
            this.btnEnd.Click += new System.EventHandler(this.btnChange_Click);
            // 
            // btnStart
            // 
            this.btnStart.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.btnStart.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnStart.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnStart.Location = new System.Drawing.Point(193, 42);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(109, 28);
            this.btnStart.TabIndex = 0;
            this.btnStart.Text = "開始";
            this.btnStart.UseVisualStyleBackColor = false;
            this.btnStart.Click += new System.EventHandler(this.btnChange_Click);
            // 
            // lblSKSMode
            // 
            this.lblSKSMode.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(176)))), ((int)(((byte)(240)))));
            this.lblSKSMode.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblSKSMode.Font = new System.Drawing.Font("MS Gothic", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSKSMode.Location = new System.Drawing.Point(52, 35);
            this.lblSKSMode.Name = "lblSKSMode";
            this.lblSKSMode.Size = new System.Drawing.Size(130, 40);
            this.lblSKSMode.TabIndex = 87;
            this.lblSKSMode.Text = "処理実行中";
            this.lblSKSMode.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // smS_Label6
            // 
            this.smS_Label6.AutoSize = true;
            this.smS_Label6.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label6.Font = new System.Drawing.Font("MS Gothic", 11F, System.Drawing.FontStyle.Bold);
            this.smS_Label6.ForeColor = System.Drawing.Color.Red;
            this.smS_Label6.Location = new System.Drawing.Point(473, 74);
            this.smS_Label6.Name = "smS_Label6";
            this.smS_Label6.Size = new System.Drawing.Size(449, 15);
            this.smS_Label6.TabIndex = 99;
            this.smS_Label6.Text = "「開始」ボタンを押す前に処理状況を確認してください。";
            // 
            // smS_Label5
            // 
            this.smS_Label5.AutoSize = true;
            this.smS_Label5.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label5.Font = new System.Drawing.Font("MS Gothic", 11F, System.Drawing.FontStyle.Bold);
            this.smS_Label5.ForeColor = System.Drawing.Color.Red;
            this.smS_Label5.Location = new System.Drawing.Point(473, 48);
            this.smS_Label5.Name = "smS_Label5";
            this.smS_Label5.Size = new System.Drawing.Size(619, 15);
            this.smS_Label5.TabIndex = 98;
            this.smS_Label5.Text = "「処理停止中」の場合、「在庫情報取得」処理を行っている可能性があります。";
            // 
            // frmPSKS0114C_1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1884, 951);
            this.KeyPreview = true;
            this.Location = new System.Drawing.Point(0, 0);
            this.Name = "frmPSKS0114C_1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "PSKS0114C_1";
            this.Load += new System.EventHandler(this.frmPSKS0114C_1_Load);
            this.panelBody.ResumeLayout(false);
            this.panelBody.PerformLayout();
            this.panelTitleLeft.ResumeLayout(false);
            this.panelTitle.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private CustomControls.SMS_Button btnEnd;
        private CustomControls.SMS_Button btnStart;
        public System.Windows.Forms.Label lblSKSMode;
        private CustomControls.SMS_Label smS_Label6;
        private CustomControls.SMS_Label smS_Label5;
    }
}