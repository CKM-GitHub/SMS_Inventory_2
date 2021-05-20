namespace SMS.Inventory
{
    partial class frmPSKS0105C
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
            this.smS_Label4 = new SMS.CustomControls.SMS_Label();
            this.smS_Label3 = new SMS.CustomControls.SMS_Label();
            this.smS_Label2 = new SMS.CustomControls.SMS_Label();
            this.lblWebMode = new System.Windows.Forms.Label();
            this.btnShow = new SMS.CustomControls.SMS_Button();
            this.btnWebStart = new SMS.CustomControls.SMS_Button();
            this.btnWebStop = new SMS.CustomControls.SMS_Button();
            this.btnInventoryStop = new SMS.CustomControls.SMS_Button();
            this.btnInventoryStart = new SMS.CustomControls.SMS_Button();
            this.lblInventoryMode = new System.Windows.Forms.Label();
            this.btnStoreStop = new SMS.CustomControls.SMS_Button();
            this.btnStoreStart = new SMS.CustomControls.SMS_Button();
            this.lblStoreMode = new System.Windows.Forms.Label();
            this.dgvPSKS0105C = new SMS.SMS_GridView();
            this.smS_Label5 = new SMS.CustomControls.SMS_Label();
            this.lblNum1 = new SMS.CustomControls.SMS_Label();
            this.colNo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colInsertDateTime = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colHanbaiShohinCD = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colVM_HANBAI_SHOHIN = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colM3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colM1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colSu = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colM2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colTekiyou = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.panelBody.SuspendLayout();
            this.panelTitleLeft.SuspendLayout();
            this.panelTitle.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvPSKS0105C)).BeginInit();
            this.SuspendLayout();
            // 
            // panelBody
            // 
            this.panelBody.Controls.Add(this.lblNum1);
            this.panelBody.Controls.Add(this.smS_Label5);
            this.panelBody.Controls.Add(this.dgvPSKS0105C);
            this.panelBody.Controls.Add(this.btnStoreStop);
            this.panelBody.Controls.Add(this.btnStoreStart);
            this.panelBody.Controls.Add(this.lblStoreMode);
            this.panelBody.Controls.Add(this.btnInventoryStop);
            this.panelBody.Controls.Add(this.btnInventoryStart);
            this.panelBody.Controls.Add(this.lblInventoryMode);
            this.panelBody.Controls.Add(this.btnWebStop);
            this.panelBody.Controls.Add(this.btnWebStart);
            this.panelBody.Controls.Add(this.btnShow);
            this.panelBody.Controls.Add(this.lblWebMode);
            this.panelBody.Controls.Add(this.smS_Label4);
            this.panelBody.Controls.Add(this.smS_Label3);
            this.panelBody.Controls.Add(this.smS_Label2);
            this.panelBody.TabIndex = 0;
            // 
            // smS_Label4
            // 
            this.smS_Label4.AutoSize = true;
            this.smS_Label4.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label4.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label4.Location = new System.Drawing.Point(1029, 26);
            this.smS_Label4.Name = "smS_Label4";
            this.smS_Label4.Size = new System.Drawing.Size(149, 12);
            this.smS_Label4.TabIndex = 29;
            this.smS_Label4.Text = "【店舗 在庫情報 取込】";
            // 
            // smS_Label3
            // 
            this.smS_Label3.AutoSize = true;
            this.smS_Label3.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label3.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label3.Location = new System.Drawing.Point(529, 27);
            this.smS_Label3.Name = "smS_Label3";
            this.smS_Label3.Size = new System.Drawing.Size(188, 12);
            this.smS_Label3.TabIndex = 28;
            this.smS_Label3.Text = "【自社在庫 入出庫情報 取込】";
            // 
            // smS_Label2
            // 
            this.smS_Label2.AutoSize = true;
            this.smS_Label2.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label2.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label2.Location = new System.Drawing.Point(22, 26);
            this.smS_Label2.Name = "smS_Label2";
            this.smS_Label2.Size = new System.Drawing.Size(183, 12);
            this.smS_Label2.TabIndex = 27;
            this.smS_Label2.Text = "【WEBストア 販売情報 取込】";
            // 
            // lblWebMode
            // 
            this.lblWebMode.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(176)))), ((int)(((byte)(240)))));
            this.lblWebMode.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblWebMode.Font = new System.Drawing.Font("MS Gothic", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblWebMode.Location = new System.Drawing.Point(48, 52);
            this.lblWebMode.Name = "lblWebMode";
            this.lblWebMode.Size = new System.Drawing.Size(130, 40);
            this.lblWebMode.TabIndex = 79;
            this.lblWebMode.Text = "処理実行中";
            this.lblWebMode.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btnShow
            // 
            this.btnShow.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.btnShow.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnShow.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnShow.Location = new System.Drawing.Point(1419, 110);
            this.btnShow.Name = "btnShow";
            this.btnShow.Size = new System.Drawing.Size(109, 28);
            this.btnShow.TabIndex = 6;
            this.btnShow.Text = "表示(F11)";
            this.btnShow.UseVisualStyleBackColor = false;
            this.btnShow.Click += new System.EventHandler(this.btnShow_Click);
            // 
            // btnWebStart
            // 
            this.btnWebStart.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.btnWebStart.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnWebStart.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnWebStart.Location = new System.Drawing.Point(189, 59);
            this.btnWebStart.Name = "btnWebStart";
            this.btnWebStart.Size = new System.Drawing.Size(109, 28);
            this.btnWebStart.TabIndex = 0;
            this.btnWebStart.Text = "開始";
            this.btnWebStart.UseVisualStyleBackColor = false;
            this.btnWebStart.Click += new System.EventHandler(this.btnWeb_Click);
            // 
            // btnWebStop
            // 
            this.btnWebStop.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.btnWebStop.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnWebStop.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnWebStop.Location = new System.Drawing.Point(298, 59);
            this.btnWebStop.Name = "btnWebStop";
            this.btnWebStop.Size = new System.Drawing.Size(109, 28);
            this.btnWebStop.TabIndex = 1;
            this.btnWebStop.Text = "停止";
            this.btnWebStop.UseVisualStyleBackColor = false;
            this.btnWebStop.Click += new System.EventHandler(this.btnWeb_Click);
            // 
            // btnInventoryStop
            // 
            this.btnInventoryStop.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.btnInventoryStop.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnInventoryStop.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnInventoryStop.Location = new System.Drawing.Point(807, 60);
            this.btnInventoryStop.Name = "btnInventoryStop";
            this.btnInventoryStop.Size = new System.Drawing.Size(109, 28);
            this.btnInventoryStop.TabIndex = 3;
            this.btnInventoryStop.Text = "停止";
            this.btnInventoryStop.UseVisualStyleBackColor = false;
            this.btnInventoryStop.Click += new System.EventHandler(this.btnInventory_Click);
            // 
            // btnInventoryStart
            // 
            this.btnInventoryStart.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.btnInventoryStart.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnInventoryStart.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnInventoryStart.Location = new System.Drawing.Point(698, 60);
            this.btnInventoryStart.Name = "btnInventoryStart";
            this.btnInventoryStart.Size = new System.Drawing.Size(109, 28);
            this.btnInventoryStart.TabIndex = 2;
            this.btnInventoryStart.Text = "開始";
            this.btnInventoryStart.UseVisualStyleBackColor = false;
            this.btnInventoryStart.Click += new System.EventHandler(this.btnInventory_Click);
            // 
            // lblInventoryMode
            // 
            this.lblInventoryMode.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(176)))), ((int)(((byte)(240)))));
            this.lblInventoryMode.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblInventoryMode.Font = new System.Drawing.Font("MS Gothic", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblInventoryMode.Location = new System.Drawing.Point(557, 53);
            this.lblInventoryMode.Name = "lblInventoryMode";
            this.lblInventoryMode.Size = new System.Drawing.Size(130, 40);
            this.lblInventoryMode.TabIndex = 83;
            this.lblInventoryMode.Text = "処理実行中";
            this.lblInventoryMode.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btnStoreStop
            // 
            this.btnStoreStop.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.btnStoreStop.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnStoreStop.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnStoreStop.Location = new System.Drawing.Point(1315, 58);
            this.btnStoreStop.Name = "btnStoreStop";
            this.btnStoreStop.Size = new System.Drawing.Size(109, 28);
            this.btnStoreStop.TabIndex = 5;
            this.btnStoreStop.Text = "停止";
            this.btnStoreStop.UseVisualStyleBackColor = false;
            this.btnStoreStop.Click += new System.EventHandler(this.btnStore_Click);
            // 
            // btnStoreStart
            // 
            this.btnStoreStart.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.btnStoreStart.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnStoreStart.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnStoreStart.Location = new System.Drawing.Point(1206, 58);
            this.btnStoreStart.Name = "btnStoreStart";
            this.btnStoreStart.Size = new System.Drawing.Size(109, 28);
            this.btnStoreStart.TabIndex = 4;
            this.btnStoreStart.Text = "開始";
            this.btnStoreStart.UseVisualStyleBackColor = false;
            this.btnStoreStart.Click += new System.EventHandler(this.btnStore_Click);
            // 
            // lblStoreMode
            // 
            this.lblStoreMode.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(176)))), ((int)(((byte)(240)))));
            this.lblStoreMode.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblStoreMode.Font = new System.Drawing.Font("MS Gothic", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblStoreMode.Location = new System.Drawing.Point(1065, 51);
            this.lblStoreMode.Name = "lblStoreMode";
            this.lblStoreMode.Size = new System.Drawing.Size(130, 40);
            this.lblStoreMode.TabIndex = 86;
            this.lblStoreMode.Text = "処理実行中";
            this.lblStoreMode.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // dgvPSKS0105C
            // 
            this.dgvPSKS0105C.AllowUserToAddRows = false;
            this.dgvPSKS0105C.AllowUserToDeleteRows = false;
            this.dgvPSKS0105C.AllowUserToResizeRows = false;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(221)))), ((int)(((byte)(235)))), ((int)(((byte)(247)))));
            this.dgvPSKS0105C.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dgvPSKS0105C.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            dataGridViewCellStyle2.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgvPSKS0105C.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle2;
            this.dgvPSKS0105C.ColumnHeadersHeight = 25;
            this.dgvPSKS0105C.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colNo,
            this.colInsertDateTime,
            this.colHanbaiShohinCD,
            this.colVM_HANBAI_SHOHIN,
            this.colM3,
            this.colM1,
            this.colSu,
            this.colM2,
            this.colTekiyou});
            dataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle6.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle6.Font = new System.Drawing.Font("MS Gothic", 9F);
            dataGridViewCellStyle6.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgvPSKS0105C.DefaultCellStyle = dataGridViewCellStyle6;
            this.dgvPSKS0105C.EnableHeadersVisualStyles = false;
            this.dgvPSKS0105C.GridColor = System.Drawing.Color.DimGray;
            this.dgvPSKS0105C.IsCellValid = false;
            this.dgvPSKS0105C.Location = new System.Drawing.Point(48, 144);
            this.dgvPSKS0105C.Name = "dgvPSKS0105C";
            this.dgvPSKS0105C.RowHeadersVisible = false;
            this.dgvPSKS0105C.Size = new System.Drawing.Size(1480, 600);
            this.dgvPSKS0105C.TabIndex = 7;
            // 
            // smS_Label5
            // 
            this.smS_Label5.AutoSize = true;
            this.smS_Label5.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label5.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label5.Location = new System.Drawing.Point(46, 129);
            this.smS_Label5.Name = "smS_Label5";
            this.smS_Label5.Size = new System.Drawing.Size(83, 12);
            this.smS_Label5.TabIndex = 90;
            this.smS_Label5.Text = "【取得履歴】";
            // 
            // lblNum1
            // 
            this.lblNum1.AutoSize = true;
            this.lblNum1.BackColor = System.Drawing.Color.Transparent;
            this.lblNum1.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.lblNum1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.lblNum1.Location = new System.Drawing.Point(155, 129);
            this.lblNum1.Name = "lblNum1";
            this.lblNum1.Size = new System.Drawing.Size(188, 12);
            this.lblNum1.TabIndex = 91;
            this.lblNum1.Text = "14日間の履歴を保持しています";
            // 
            // colNo
            // 
            this.colNo.DataPropertyName = "No";
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            this.colNo.DefaultCellStyle = dataGridViewCellStyle3;
            this.colNo.HeaderText = "No";
            this.colNo.MaxInputLength = 4;
            this.colNo.Name = "colNo";
            this.colNo.Width = 50;
            // 
            // colInsertDateTime
            // 
            this.colInsertDateTime.DataPropertyName = "InsertDateTime";
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this.colInsertDateTime.DefaultCellStyle = dataGridViewCellStyle4;
            this.colInsertDateTime.HeaderText = "日時";
            this.colInsertDateTime.MaxInputLength = 12;
            this.colInsertDateTime.Name = "colInsertDateTime";
            this.colInsertDateTime.Width = 150;
            // 
            // colHanbaiShohinCD
            // 
            this.colHanbaiShohinCD.DataPropertyName = "HanbaiShohinCD";
            this.colHanbaiShohinCD.HeaderText = "SKUCD";
            this.colHanbaiShohinCD.MaxInputLength = 30;
            this.colHanbaiShohinCD.Name = "colHanbaiShohinCD";
            this.colHanbaiShohinCD.Width = 160;
            // 
            // colVM_HANBAI_SHOHIN
            // 
            this.colVM_HANBAI_SHOHIN.DataPropertyName = "VM_HANBAI_SHOHIN";
            this.colVM_HANBAI_SHOHIN.HeaderText = "商品名";
            this.colVM_HANBAI_SHOHIN.MaxInputLength = 80;
            this.colVM_HANBAI_SHOHIN.Name = "colVM_HANBAI_SHOHIN";
            this.colVM_HANBAI_SHOHIN.Width = 530;
            // 
            // colM3
            // 
            this.colM3.DataPropertyName = "M3";
            this.colM3.HeaderText = "倉庫";
            this.colM3.MaxInputLength = 20;
            this.colM3.Name = "colM3";
            this.colM3.Width = 150;
            // 
            // colM1
            // 
            this.colM1.DataPropertyName = "M1";
            this.colM1.HeaderText = "増減理由";
            this.colM1.MaxInputLength = 20;
            this.colM1.Name = "colM1";
            this.colM1.Width = 120;
            // 
            // colSu
            // 
            this.colSu.DataPropertyName = "Su";
            dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            this.colSu.DefaultCellStyle = dataGridViewCellStyle5;
            this.colSu.HeaderText = "増減数";
            this.colSu.MaxInputLength = 4;
            this.colSu.Name = "colSu";
            this.colSu.Width = 60;
            // 
            // colM2
            // 
            this.colM2.DataPropertyName = "M2";
            this.colM2.HeaderText = "店舗";
            this.colM2.MaxInputLength = 20;
            this.colM2.Name = "colM2";
            this.colM2.Width = 120;
            // 
            // colTekiyou
            // 
            this.colTekiyou.DataPropertyName = "Tekiyou";
            this.colTekiyou.HeaderText = "摘要";
            this.colTekiyou.MaxInputLength = 50;
            this.colTekiyou.Name = "colTekiyou";
            this.colTekiyou.Width = 120;
            // 
            // frmPSKS0105C
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1884, 951);
            this.KeyPreview = true;
            this.Location = new System.Drawing.Point(0, 0);
            this.Name = "frmPSKS0105C";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "PSKS0105C";
            this.Load += new System.EventHandler(this.frmPSKS0105C_Load);
            this.panelBody.ResumeLayout(false);
            this.panelBody.PerformLayout();
            this.panelTitleLeft.ResumeLayout(false);
            this.panelTitle.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvPSKS0105C)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private CustomControls.SMS_Label smS_Label4;
        private CustomControls.SMS_Label smS_Label3;
        private CustomControls.SMS_Label smS_Label2;
        public System.Windows.Forms.Label lblWebMode;
        private CustomControls.SMS_Label smS_Label5;
        private SMS_GridView dgvPSKS0105C;
        private CustomControls.SMS_Button btnStoreStop;
        private CustomControls.SMS_Button btnStoreStart;
        public System.Windows.Forms.Label lblStoreMode;
        private CustomControls.SMS_Button btnInventoryStop;
        private CustomControls.SMS_Button btnInventoryStart;
        public System.Windows.Forms.Label lblInventoryMode;
        private CustomControls.SMS_Button btnWebStop;
        private CustomControls.SMS_Button btnWebStart;
        private CustomControls.SMS_Button btnShow;
        private CustomControls.SMS_Label lblNum1;
        private System.Windows.Forms.DataGridViewTextBoxColumn colNo;
        private System.Windows.Forms.DataGridViewTextBoxColumn colInsertDateTime;
        private System.Windows.Forms.DataGridViewTextBoxColumn colHanbaiShohinCD;
        private System.Windows.Forms.DataGridViewTextBoxColumn colVM_HANBAI_SHOHIN;
        private System.Windows.Forms.DataGridViewTextBoxColumn colM3;
        private System.Windows.Forms.DataGridViewTextBoxColumn colM1;
        private System.Windows.Forms.DataGridViewTextBoxColumn colSu;
        private System.Windows.Forms.DataGridViewTextBoxColumn colM2;
        private System.Windows.Forms.DataGridViewTextBoxColumn colTekiyou;
    }
}