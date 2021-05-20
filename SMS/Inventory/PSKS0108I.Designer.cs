namespace SMS.Inventory
{
    partial class frmPSKS0108I
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
            this.btnStage4 = new SMS.CustomControls.SMS_Button();
            this.btnStage3Save = new SMS.CustomControls.SMS_Button();
            this.smS_Label4 = new SMS.CustomControls.SMS_Label();
            this.btnStage3 = new SMS.CustomControls.SMS_Button();
            this.btnWebSave = new SMS.CustomControls.SMS_Button();
            this.smS_Label2 = new SMS.CustomControls.SMS_Label();
            this.btnWeb = new SMS.CustomControls.SMS_Button();
            this.txtStartTime = new SMS.CustomControls.DateTimeTextBox();
            this.txtEndTime = new SMS.CustomControls.DateTimeTextBox();
            this.panel2 = new System.Windows.Forms.Panel();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel3 = new System.Windows.Forms.Panel();
            this.lblStage5 = new System.Windows.Forms.Label();
            this.lblStage6 = new System.Windows.Forms.Label();
            this.lblStage7 = new System.Windows.Forms.Label();
            this.lblStage2 = new System.Windows.Forms.Label();
            this.panelBody.SuspendLayout();
            this.panelTitleLeft.SuspendLayout();
            this.panelTitle.SuspendLayout();
            this.SuspendLayout();
            // 
            // panelBody
            // 
            this.panelBody.Controls.Add(this.lblStage2);
            this.panelBody.Controls.Add(this.lblStage7);
            this.panelBody.Controls.Add(this.lblStage6);
            this.panelBody.Controls.Add(this.lblStage5);
            this.panelBody.Controls.Add(this.panel3);
            this.panelBody.Controls.Add(this.panel1);
            this.panelBody.Controls.Add(this.txtEndTime);
            this.panelBody.Controls.Add(this.txtStartTime);
            this.panelBody.Controls.Add(this.btnStage4);
            this.panelBody.Controls.Add(this.btnStage3Save);
            this.panelBody.Controls.Add(this.smS_Label4);
            this.panelBody.Controls.Add(this.btnStage3);
            this.panelBody.Controls.Add(this.btnWebSave);
            this.panelBody.Controls.Add(this.panel2);
            this.panelBody.Controls.Add(this.smS_Label2);
            this.panelBody.Controls.Add(this.btnWeb);
            this.panelBody.Location = new System.Drawing.Point(0, 36);
            this.panelBody.Size = new System.Drawing.Size(1884, 885);
            this.panelBody.TabIndex = 0;
            // 
            // panelTitleLeft
            // 
            this.panelTitleLeft.Size = new System.Drawing.Size(644, 36);
            // 
            // panelTitle
            // 
            this.panelTitle.Size = new System.Drawing.Size(1884, 36);
            // 
            // btnStage4
            // 
            this.btnStage4.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.btnStage4.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnStage4.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnStage4.Font = new System.Drawing.Font("MS Gothic", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnStage4.Location = new System.Drawing.Point(36, 363);
            this.btnStage4.Name = "btnStage4";
            this.btnStage4.Size = new System.Drawing.Size(1000, 40);
            this.btnStage4.TabIndex = 28;
            this.btnStage4.Text = "④「すご楽」最新在庫情報の取得開始";
            this.btnStage4.UseVisualStyleBackColor = false;
            this.btnStage4.Click += new System.EventHandler(this.btnStage4_Click);
            // 
            // btnStage3Save
            // 
            this.btnStage3Save.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.btnStage3Save.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnStage3Save.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnStage3Save.Location = new System.Drawing.Point(861, 306);
            this.btnStage3Save.Name = "btnStage3Save";
            this.btnStage3Save.Size = new System.Drawing.Size(100, 30);
            this.btnStage3Save.TabIndex = 3;
            this.btnStage3Save.Text = "登録";
            this.btnStage3Save.UseVisualStyleBackColor = false;
            this.btnStage3Save.Click += new System.EventHandler(this.btnStage3Save_Click);
            // 
            // smS_Label4
            // 
            this.smS_Label4.AutoSize = true;
            this.smS_Label4.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label4.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label4.Location = new System.Drawing.Point(598, 315);
            this.smS_Label4.Name = "smS_Label4";
            this.smS_Label4.Size = new System.Drawing.Size(200, 12);
            this.smS_Label4.TabIndex = 24;
            this.smS_Label4.Text = "押し忘れた場合、その時刻を入力";
            // 
            // btnStage3
            // 
            this.btnStage3.BackColor = System.Drawing.Color.Silver;
            this.btnStage3.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnStage3.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnStage3.Font = new System.Drawing.Font("MS Gothic", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnStage3.Location = new System.Drawing.Point(36, 251);
            this.btnStage3.Name = "btnStage3";
            this.btnStage3.Size = new System.Drawing.Size(1000, 40);
            this.btnStage3.TabIndex = 23;
            this.btnStage3.Text = "③「すご楽」への取込が全ストア分終了直後に、このボタンを押す";
            this.btnStage3.UseVisualStyleBackColor = false;
            this.btnStage3.Click += new System.EventHandler(this.btnStage3_Click);
            // 
            // btnWebSave
            // 
            this.btnWebSave.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.btnWebSave.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnWebSave.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnWebSave.Location = new System.Drawing.Point(861, 79);
            this.btnWebSave.Name = "btnWebSave";
            this.btnWebSave.Size = new System.Drawing.Size(100, 30);
            this.btnWebSave.TabIndex = 1;
            this.btnWebSave.Text = "登録";
            this.btnWebSave.UseVisualStyleBackColor = false;
            this.btnWebSave.Click += new System.EventHandler(this.btnWebSave_Click);
            // 
            // smS_Label2
            // 
            this.smS_Label2.AutoSize = true;
            this.smS_Label2.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label2.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label2.Location = new System.Drawing.Point(598, 89);
            this.smS_Label2.Name = "smS_Label2";
            this.smS_Label2.Size = new System.Drawing.Size(200, 12);
            this.smS_Label2.TabIndex = 17;
            this.smS_Label2.Text = "押し忘れた場合、その時刻を入力";
            // 
            // btnWeb
            // 
            this.btnWeb.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.btnWeb.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnWeb.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnWeb.Font = new System.Drawing.Font("MS Gothic", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnWeb.Location = new System.Drawing.Point(36, 27);
            this.btnWeb.Name = "btnWeb";
            this.btnWeb.Size = new System.Drawing.Size(1000, 40);
            this.btnWeb.TabIndex = 15;
            this.btnWeb.Text = "① WEBストア受注情報CSVの「すご楽」への取込の開始直前、このボタンを押す";
            this.btnWeb.UseVisualStyleBackColor = false;
            this.btnWeb.Click += new System.EventHandler(this.btnWeb_Click);
            // 
            // txtStartTime
            // 
            this.txtStartTime.Location = new System.Drawing.Point(803, 85);
            this.txtStartTime.Mask = "00:00";
            this.txtStartTime.MaskFormat = "HH:mm";
            this.txtStartTime.Name = "txtStartTime";
            this.txtStartTime.PromptChar = ' ';
            this.txtStartTime.Size = new System.Drawing.Size(52, 19);
            this.txtStartTime.TabIndex = 0;
            // 
            // txtEndTime
            // 
            this.txtEndTime.Location = new System.Drawing.Point(804, 312);
            this.txtEndTime.Mask = "00:00";
            this.txtEndTime.MaskFormat = "HH:mm";
            this.txtEndTime.Name = "txtEndTime";
            this.txtEndTime.PromptChar = ' ';
            this.txtEndTime.Size = new System.Drawing.Size(52, 19);
            this.txtEndTime.TabIndex = 2;
            // 
            // panel2
            // 
            this.panel2.BackgroundImage = global::SMS.Properties.Resources.downarrow;
            this.panel2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.panel2.Location = new System.Drawing.Point(501, 73);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(80, 60);
            this.panel2.TabIndex = 18;
            // 
            // panel1
            // 
            this.panel1.BackgroundImage = global::SMS.Properties.Resources.downarrow;
            this.panel1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.panel1.Location = new System.Drawing.Point(501, 185);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(80, 60);
            this.panel1.TabIndex = 19;
            // 
            // panel3
            // 
            this.panel3.BackgroundImage = global::SMS.Properties.Resources.downarrow;
            this.panel3.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.panel3.Location = new System.Drawing.Point(501, 297);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(80, 60);
            this.panel3.TabIndex = 33;
            // 
            // lblStage5
            // 
            this.lblStage5.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(242)))), ((int)(((byte)(204)))));
            this.lblStage5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblStage5.Font = new System.Drawing.Font("MS Gothic", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblStage5.Location = new System.Drawing.Point(36, 417);
            this.lblStage5.Name = "lblStage5";
            this.lblStage5.Size = new System.Drawing.Size(1000, 40);
            this.lblStage5.TabIndex = 46;
            this.lblStage5.Text = "⑤「すご楽」最新在庫情報を取込中";
            this.lblStage5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lblStage6
            // 
            this.lblStage6.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(242)))), ((int)(((byte)(204)))));
            this.lblStage6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblStage6.Font = new System.Drawing.Font("MS Gothic", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblStage6.Location = new System.Drawing.Point(39, 471);
            this.lblStage6.Name = "lblStage6";
            this.lblStage6.Size = new System.Drawing.Size(1000, 40);
            this.lblStage6.TabIndex = 47;
            this.lblStage6.Text = "⑥この処理実行中(①～③)にWEBストアで販売された結果を在庫に反映開始";
            this.lblStage6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lblStage7
            // 
            this.lblStage7.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(242)))), ((int)(((byte)(204)))));
            this.lblStage7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblStage7.Font = new System.Drawing.Font("MS Gothic", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblStage7.Location = new System.Drawing.Point(39, 524);
            this.lblStage7.Name = "lblStage7";
            this.lblStage7.Size = new System.Drawing.Size(1000, 40);
            this.lblStage7.TabIndex = 48;
            this.lblStage7.Text = "⑦全ての処理完了（次回の在庫SKS連携処理でSKSに連携する）";
            this.lblStage7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lblStage2
            // 
            this.lblStage2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(242)))), ((int)(((byte)(204)))));
            this.lblStage2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblStage2.Font = new System.Drawing.Font("MS Gothic", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblStage2.Location = new System.Drawing.Point(36, 141);
            this.lblStage2.Name = "lblStage2";
            this.lblStage2.Size = new System.Drawing.Size(1000, 40);
            this.lblStage2.TabIndex = 49;
            this.lblStage2.Text = "②「すご楽」で従来どおり受注情報CSVを取込する";
            this.lblStage2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // frmPSKS0108I
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1884, 951);
            this.F5_Text = "ｷｬﾝｾﾙ(F5)";
            this.KeyPreview = true;
            this.Location = new System.Drawing.Point(0, 0);
            this.Name = "frmPSKS0108I";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "PSKS0108I";
            this.Load += new System.EventHandler(this.frmPSKS0108I_Load);
            this.panelBody.ResumeLayout(false);
            this.panelBody.PerformLayout();
            this.panelTitleLeft.ResumeLayout(false);
            this.panelTitle.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private CustomControls.SMS_Button btnStage4;
        private CustomControls.SMS_Button btnStage3Save;
        private CustomControls.SMS_Label smS_Label4;
        private CustomControls.SMS_Button btnStage3;
        private CustomControls.SMS_Button btnWebSave;
        private System.Windows.Forms.Panel panel2;
        private CustomControls.SMS_Label smS_Label2;
        private CustomControls.SMS_Button btnWeb;
        private CustomControls.DateTimeTextBox txtStartTime;
        private CustomControls.DateTimeTextBox txtEndTime;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label lblStage7;
        private System.Windows.Forms.Label lblStage6;
        private System.Windows.Forms.Label lblStage5;
        private System.Windows.Forms.Label lblStage2;
    }
}