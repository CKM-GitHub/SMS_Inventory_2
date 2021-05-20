namespace SMS.Order
{
    partial class F0004
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
            this.cboprocessing = new System.Windows.Forms.ComboBox();
            this.smS_Label2 = new SMS.CustomControls.SMS_Label();
            this.panelBody.SuspendLayout();
            this.panelTitleLeft.SuspendLayout();
            this.panelTitle.SuspendLayout();
            this.SuspendLayout();
            // 
            // panelBody
            // 
            this.panelBody.Controls.Add(this.cboprocessing);
            this.panelBody.Controls.Add(this.smS_Label2);
            this.panelBody.TabIndex = 0;
            // 
            // cboprocessing
            // 
            this.cboprocessing.FormattingEnabled = true;
            this.cboprocessing.Location = new System.Drawing.Point(95, 27);
            this.cboprocessing.Name = "cboprocessing";
            this.cboprocessing.Size = new System.Drawing.Size(200, 20);
            this.cboprocessing.TabIndex = 0;
            // 
            // smS_Label2
            // 
            this.smS_Label2.AutoSize = true;
            this.smS_Label2.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label2.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label2.Location = new System.Drawing.Point(58, 31);
            this.smS_Label2.Name = "smS_Label2";
            this.smS_Label2.Size = new System.Drawing.Size(31, 12);
            this.smS_Label2.TabIndex = 2;
            this.smS_Label2.Text = "処理";
            // 
            // F0004
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1884, 951);
            this.F12_Text = "開始(F12)";
            this.KeyPreview = true;
            this.Location = new System.Drawing.Point(0, 0);
            this.Name = "F0004";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "F0004";
            this.Load += new System.EventHandler(this.F0004_Load);
            this.panelBody.ResumeLayout(false);
            this.panelBody.PerformLayout();
            this.panelTitleLeft.ResumeLayout(false);
            this.panelTitle.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ComboBox cboprocessing;
        private CustomControls.SMS_Label smS_Label2;
    }
}