namespace SMS.Inventory
{
    partial class frmPSKS0119M
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle7 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle8 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle9 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle10 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            this.smS_Label2 = new SMS.CustomControls.SMS_Label();
            this.txtFilePath = new SMS.CustomControls.SMS_TextBox();
            this.btnBrowse = new SMS.CustomControls.SMS_Button();
            this.btnImport = new SMS.CustomControls.SMS_Button();
            this.btnDisplay = new SMS.CustomControls.SMS_Button();
            this.smS_Label3 = new SMS.CustomControls.SMS_Label();
            this.smS_Label4 = new SMS.CustomControls.SMS_Label();
            this.smS_Label5 = new SMS.CustomControls.SMS_Label();
            this.smS_Label6 = new SMS.CustomControls.SMS_Label();
            this.cboTaisho = new SMS.CustomControls.SMS_ComboBox();
            this.cboShiiresaki = new SMS.CustomControls.SMS_ComboBox();
            this.cboBrand = new SMS.CustomControls.SMS_ComboBox();
            this.btnAllCheck = new SMS.CustomControls.SMS_Button();
            this.btnAllUncheck = new SMS.CustomControls.SMS_Button();
            this.smS_Label7 = new SMS.CustomControls.SMS_Label();
            this.txtRefNo = new SMS.CustomControls.SMS_TextBox();
            this.smS_Label8 = new SMS.CustomControls.SMS_Label();
            this.cboZaikoKeisan = new SMS.CustomControls.SMS_ComboBox();
            this.btnItemSearch = new SMS.CustomControls.SMS_Button();
            this.btnSetting = new SMS.CustomControls.SMS_Button();
            this.gdvPSKS0119M = new SMS.SMS_GridView();
            this.smS_GridView1 = new SMS.SMS_GridView();
            this.gdv = new SMS.SMS_GridView();
            this.smS_Label9 = new SMS.CustomControls.SMS_Label();
            this.cboTaisho1 = new SMS.CustomControls.SMS_ComboBox();
            this.colNo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colChk = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.colTaisho = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colSKUCD = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colJANCD = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colMakerCD = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colItemName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colColorName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colSizeName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colRefNo = new SMS.CustomControls.DataGridViewDecimalColumn();
            this.colZaikoKeisan = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.colShiiresakiName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colBrandName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colInstructionNo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colRemark = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colAdminCD = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.panelBody.SuspendLayout();
            this.panelTitleLeft.SuspendLayout();
            this.panelTitle.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gdvPSKS0119M)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.smS_GridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gdv)).BeginInit();
            this.SuspendLayout();
            // 
            // panelBody
            // 
            this.panelBody.Controls.Add(this.cboTaisho1);
            this.panelBody.Controls.Add(this.smS_Label9);
            this.panelBody.Controls.Add(this.gdvPSKS0119M);
            this.panelBody.Controls.Add(this.btnItemSearch);
            this.panelBody.Controls.Add(this.btnSetting);
            this.panelBody.Controls.Add(this.cboZaikoKeisan);
            this.panelBody.Controls.Add(this.smS_Label8);
            this.panelBody.Controls.Add(this.txtRefNo);
            this.panelBody.Controls.Add(this.smS_Label7);
            this.panelBody.Controls.Add(this.btnAllUncheck);
            this.panelBody.Controls.Add(this.btnAllCheck);
            this.panelBody.Controls.Add(this.cboBrand);
            this.panelBody.Controls.Add(this.cboShiiresaki);
            this.panelBody.Controls.Add(this.cboTaisho);
            this.panelBody.Controls.Add(this.smS_Label6);
            this.panelBody.Controls.Add(this.smS_Label5);
            this.panelBody.Controls.Add(this.smS_Label4);
            this.panelBody.Controls.Add(this.smS_Label3);
            this.panelBody.Controls.Add(this.btnDisplay);
            this.panelBody.Controls.Add(this.btnImport);
            this.panelBody.Controls.Add(this.btnBrowse);
            this.panelBody.Controls.Add(this.txtFilePath);
            this.panelBody.Controls.Add(this.smS_Label2);
            this.panelBody.TabIndex = 0;
            // 
            // smS_Label2
            // 
            this.smS_Label2.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.smS_Label2.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label2.Location = new System.Drawing.Point(15, 105);
            this.smS_Label2.Name = "smS_Label2";
            this.smS_Label2.Size = new System.Drawing.Size(1850, 1);
            this.smS_Label2.TabIndex = 0;
            // 
            // txtFilePath
            // 
            this.txtFilePath.Decimalplace = ((byte)(0));
            this.txtFilePath.Location = new System.Drawing.Point(66, 30);
            this.txtFilePath.Name = "txtFilePath";
            this.txtFilePath.Size = new System.Drawing.Size(450, 19);
            this.txtFilePath.TabIndex = 1;
            this.txtFilePath.txt_Type = SMS.CustomControls.SMS_TextBox.Type.Normal;
            this.txtFilePath.UseMinus = false;
            this.txtFilePath.UseThousandSeparator = true;
            // 
            // btnBrowse
            // 
            this.btnBrowse.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.btnBrowse.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnBrowse.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnBrowse.Location = new System.Drawing.Point(516, 30);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(90, 19);
            this.btnBrowse.TabIndex = 2;
            this.btnBrowse.Text = "ファイル選択";
            this.btnBrowse.UseVisualStyleBackColor = false;
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // btnImport
            // 
            this.btnImport.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.btnImport.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnImport.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnImport.Location = new System.Drawing.Point(1758, 69);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(100, 30);
            this.btnImport.TabIndex = 3;
            this.btnImport.Text = "取込";
            this.btnImport.UseVisualStyleBackColor = false;
            this.btnImport.Click += new System.EventHandler(this.btnImport_Click);
            // 
            // btnDisplay
            // 
            this.btnDisplay.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.btnDisplay.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnDisplay.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnDisplay.Location = new System.Drawing.Point(1761, 142);
            this.btnDisplay.Name = "btnDisplay";
            this.btnDisplay.Size = new System.Drawing.Size(100, 30);
            this.btnDisplay.TabIndex = 7;
            this.btnDisplay.Text = "表示（F10）";
            this.btnDisplay.UseVisualStyleBackColor = false;
            this.btnDisplay.Click += new System.EventHandler(this.btnDisplay_Click);
            // 
            // smS_Label3
            // 
            this.smS_Label3.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.smS_Label3.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label3.Location = new System.Drawing.Point(15, 181);
            this.smS_Label3.Name = "smS_Label3";
            this.smS_Label3.Size = new System.Drawing.Size(1850, 2);
            this.smS_Label3.TabIndex = 5;
            // 
            // smS_Label4
            // 
            this.smS_Label4.AutoSize = true;
            this.smS_Label4.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label4.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label4.Location = new System.Drawing.Point(82, 116);
            this.smS_Label4.Name = "smS_Label4";
            this.smS_Label4.Size = new System.Drawing.Size(31, 12);
            this.smS_Label4.TabIndex = 6;
            this.smS_Label4.Text = "対象";
            // 
            // smS_Label5
            // 
            this.smS_Label5.AutoSize = true;
            this.smS_Label5.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label5.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label5.Location = new System.Drawing.Point(197, 117);
            this.smS_Label5.Name = "smS_Label5";
            this.smS_Label5.Size = new System.Drawing.Size(44, 12);
            this.smS_Label5.TabIndex = 7;
            this.smS_Label5.Text = "仕入先";
            // 
            // smS_Label6
            // 
            this.smS_Label6.AutoSize = true;
            this.smS_Label6.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label6.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label6.Location = new System.Drawing.Point(416, 117);
            this.smS_Label6.Name = "smS_Label6";
            this.smS_Label6.Size = new System.Drawing.Size(57, 12);
            this.smS_Label6.TabIndex = 8;
            this.smS_Label6.Text = "ブランド";
            // 
            // cboTaisho
            // 
            this.cboTaisho.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.cboTaisho.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource;
            this.cboTaisho.Cbo_Type = SMS.CustomControls.SMS_ComboBox.cbo_Type.Taisho;
            this.cboTaisho.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cboTaisho.FormattingEnabled = true;
            this.cboTaisho.Location = new System.Drawing.Point(116, 113);
            this.cboTaisho.Name = "cboTaisho";
            this.cboTaisho.Size = new System.Drawing.Size(60, 20);
            this.cboTaisho.TabIndex = 4;
            // 
            // cboShiiresaki
            // 
            this.cboShiiresaki.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.cboShiiresaki.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource;
            this.cboShiiresaki.Cbo_Type = SMS.CustomControls.SMS_ComboBox.cbo_Type.M_SHIIRESAKI;
            this.cboShiiresaki.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cboShiiresaki.FormattingEnabled = true;
            this.cboShiiresaki.Location = new System.Drawing.Point(244, 114);
            this.cboShiiresaki.Name = "cboShiiresaki";
            this.cboShiiresaki.Size = new System.Drawing.Size(150, 20);
            this.cboShiiresaki.TabIndex = 5;
            // 
            // cboBrand
            // 
            this.cboBrand.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.cboBrand.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource;
            this.cboBrand.Cbo_Type = SMS.CustomControls.SMS_ComboBox.cbo_Type.m_brand;
            this.cboBrand.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cboBrand.FormattingEnabled = true;
            this.cboBrand.Location = new System.Drawing.Point(476, 114);
            this.cboBrand.Name = "cboBrand";
            this.cboBrand.Size = new System.Drawing.Size(150, 20);
            this.cboBrand.TabIndex = 6;
            // 
            // btnAllCheck
            // 
            this.btnAllCheck.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.btnAllCheck.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnAllCheck.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnAllCheck.Location = new System.Drawing.Point(34, 190);
            this.btnAllCheck.Name = "btnAllCheck";
            this.btnAllCheck.Size = new System.Drawing.Size(100, 30);
            this.btnAllCheck.TabIndex = 8;
            this.btnAllCheck.Text = "全選択";
            this.btnAllCheck.UseVisualStyleBackColor = false;
            this.btnAllCheck.Click += new System.EventHandler(this.btnAllCheck_Click);
            // 
            // btnAllUncheck
            // 
            this.btnAllUncheck.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.btnAllUncheck.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnAllUncheck.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnAllUncheck.Location = new System.Drawing.Point(134, 190);
            this.btnAllUncheck.Name = "btnAllUncheck";
            this.btnAllUncheck.Size = new System.Drawing.Size(100, 30);
            this.btnAllUncheck.TabIndex = 9;
            this.btnAllUncheck.Text = "全解除";
            this.btnAllUncheck.UseVisualStyleBackColor = false;
            this.btnAllUncheck.Click += new System.EventHandler(this.btnAllUncheck_Click);
            // 
            // smS_Label7
            // 
            this.smS_Label7.AutoSize = true;
            this.smS_Label7.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label7.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label7.Location = new System.Drawing.Point(710, 192);
            this.smS_Label7.Name = "smS_Label7";
            this.smS_Label7.Size = new System.Drawing.Size(44, 12);
            this.smS_Label7.TabIndex = 14;
            this.smS_Label7.Text = "基準数";
            // 
            // txtRefNo
            // 
            this.txtRefNo.Decimalplace = ((byte)(0));
            this.txtRefNo.Location = new System.Drawing.Point(757, 190);
            this.txtRefNo.Name = "txtRefNo";
            this.txtRefNo.Size = new System.Drawing.Size(70, 19);
            this.txtRefNo.TabIndex = 10;
            this.txtRefNo.txt_Type = SMS.CustomControls.SMS_TextBox.Type.Normal;
            this.txtRefNo.UseMinus = false;
            this.txtRefNo.UseThousandSeparator = true;
            // 
            // smS_Label8
            // 
            this.smS_Label8.AutoSize = true;
            this.smS_Label8.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label8.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label8.Location = new System.Drawing.Point(847, 192);
            this.smS_Label8.Name = "smS_Label8";
            this.smS_Label8.Size = new System.Drawing.Size(57, 12);
            this.smS_Label8.TabIndex = 16;
            this.smS_Label8.Text = "在庫計算";
            // 
            // cboZaikoKeisan
            // 
            this.cboZaikoKeisan.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.cboZaikoKeisan.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource;
            this.cboZaikoKeisan.Cbo_Type = SMS.CustomControls.SMS_ComboBox.cbo_Type.ZaikoKeisan;
            this.cboZaikoKeisan.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cboZaikoKeisan.FormattingEnabled = true;
            this.cboZaikoKeisan.Location = new System.Drawing.Point(907, 189);
            this.cboZaikoKeisan.Name = "cboZaikoKeisan";
            this.cboZaikoKeisan.Size = new System.Drawing.Size(185, 20);
            this.cboZaikoKeisan.TabIndex = 11;
            // 
            // btnItemSearch
            // 
            this.btnItemSearch.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.btnItemSearch.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnItemSearch.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnItemSearch.Location = new System.Drawing.Point(1762, 189);
            this.btnItemSearch.Name = "btnItemSearch";
            this.btnItemSearch.Size = new System.Drawing.Size(100, 30);
            this.btnItemSearch.TabIndex = 14;
            this.btnItemSearch.Text = "対象商品追加";
            this.btnItemSearch.UseVisualStyleBackColor = false;
            this.btnItemSearch.Click += new System.EventHandler(this.btnItemSearch_Click);
            // 
            // btnSetting
            // 
            this.btnSetting.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.btnSetting.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnSetting.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnSetting.Location = new System.Drawing.Point(1640, 189);
            this.btnSetting.Name = "btnSetting";
            this.btnSetting.Size = new System.Drawing.Size(100, 30);
            this.btnSetting.TabIndex = 13;
            this.btnSetting.Text = "設定";
            this.btnSetting.UseVisualStyleBackColor = false;
            this.btnSetting.Click += new System.EventHandler(this.btnSetting_Click);
            // 
            // gdvPSKS0119M
            // 
            this.gdvPSKS0119M.AllowUserToAddRows = false;
            this.gdvPSKS0119M.AllowUserToDeleteRows = false;
            this.gdvPSKS0119M.AllowUserToResizeRows = false;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(221)))), ((int)(((byte)(235)))), ((int)(((byte)(247)))));
            this.gdvPSKS0119M.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.gdvPSKS0119M.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            dataGridViewCellStyle2.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.gdvPSKS0119M.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle2;
            this.gdvPSKS0119M.ColumnHeadersHeight = 25;
            this.gdvPSKS0119M.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colNo,
            this.colChk,
            this.colTaisho,
            this.colSKUCD,
            this.colJANCD,
            this.colMakerCD,
            this.colItemName,
            this.colColorName,
            this.colSizeName,
            this.colRefNo,
            this.colZaikoKeisan,
            this.colShiiresakiName,
            this.colBrandName,
            this.colInstructionNo,
            this.colRemark,
            this.colAdminCD});
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle4.Font = new System.Drawing.Font("MS Gothic", 9F);
            dataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.gdvPSKS0119M.DefaultCellStyle = dataGridViewCellStyle4;
            this.gdvPSKS0119M.EnableHeadersVisualStyles = false;
            this.gdvPSKS0119M.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            this.gdvPSKS0119M.IsCellValid = false;
            this.gdvPSKS0119M.Location = new System.Drawing.Point(16, 228);
            this.gdvPSKS0119M.Name = "gdvPSKS0119M";
            this.gdvPSKS0119M.RowHeadersVisible = false;
            this.gdvPSKS0119M.Size = new System.Drawing.Size(1850, 600);
            this.gdvPSKS0119M.TabIndex = 20;
            this.gdvPSKS0119M.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.gdvPSKS0119M_CellContentClick);
            this.gdvPSKS0119M.CellEnter += new System.Windows.Forms.DataGridViewCellEventHandler(this.gdvPSKS0119M_CellEnter);
            this.gdvPSKS0119M.CellFormatting += new System.Windows.Forms.DataGridViewCellFormattingEventHandler(this.gdvPSKS0119M_CellFormatting);
            // 
            // smS_GridView1
            // 
            this.smS_GridView1.AllowUserToAddRows = false;
            this.smS_GridView1.AllowUserToDeleteRows = false;
            this.smS_GridView1.AllowUserToResizeRows = false;
            dataGridViewCellStyle5.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(221)))), ((int)(((byte)(235)))), ((int)(((byte)(247)))));
            this.smS_GridView1.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle5;
            this.smS_GridView1.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            dataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle6.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            dataGridViewCellStyle6.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            dataGridViewCellStyle6.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.smS_GridView1.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle6;
            this.smS_GridView1.ColumnHeadersHeight = 25;
            dataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle7.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle7.Font = new System.Drawing.Font("MS Gothic", 9F);
            dataGridViewCellStyle7.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle7.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle7.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.smS_GridView1.DefaultCellStyle = dataGridViewCellStyle7;
            this.smS_GridView1.EnableHeadersVisualStyles = false;
            this.smS_GridView1.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            this.smS_GridView1.IsCellValid = false;
            this.smS_GridView1.Location = new System.Drawing.Point(16, 250);
            this.smS_GridView1.Name = "smS_GridView1";
            this.smS_GridView1.RowHeadersVisible = false;
            this.smS_GridView1.Size = new System.Drawing.Size(1700, 500);
            this.smS_GridView1.TabIndex = 20;
            // 
            // gdv
            // 
            this.gdv.AllowUserToAddRows = false;
            this.gdv.AllowUserToDeleteRows = false;
            this.gdv.AllowUserToResizeRows = false;
            dataGridViewCellStyle8.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(221)))), ((int)(((byte)(235)))), ((int)(((byte)(247)))));
            this.gdv.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle8;
            this.gdv.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            dataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle9.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            dataGridViewCellStyle9.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            dataGridViewCellStyle9.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle9.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle9.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle9.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.gdv.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle9;
            this.gdv.ColumnHeadersHeight = 25;
            dataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle10.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle10.Font = new System.Drawing.Font("MS Gothic", 9F);
            dataGridViewCellStyle10.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle10.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle10.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle10.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.gdv.DefaultCellStyle = dataGridViewCellStyle10;
            this.gdv.EnableHeadersVisualStyles = false;
            this.gdv.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            this.gdv.IsCellValid = false;
            this.gdv.Location = new System.Drawing.Point(16, 250);
            this.gdv.Name = "gdv";
            this.gdv.RowHeadersVisible = false;
            this.gdv.Size = new System.Drawing.Size(1700, 500);
            this.gdv.TabIndex = 20;
            // 
            // smS_Label9
            // 
            this.smS_Label9.AutoSize = true;
            this.smS_Label9.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label9.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label9.Location = new System.Drawing.Point(1124, 192);
            this.smS_Label9.Name = "smS_Label9";
            this.smS_Label9.Size = new System.Drawing.Size(31, 12);
            this.smS_Label9.TabIndex = 21;
            this.smS_Label9.Text = "対象";
            // 
            // cboTaisho1
            // 
            this.cboTaisho1.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.cboTaisho1.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource;
            this.cboTaisho1.Cbo_Type = SMS.CustomControls.SMS_ComboBox.cbo_Type.Taisho;
            this.cboTaisho1.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cboTaisho1.FormattingEnabled = true;
            this.cboTaisho1.Location = new System.Drawing.Point(1158, 189);
            this.cboTaisho1.Name = "cboTaisho1";
            this.cboTaisho1.Size = new System.Drawing.Size(60, 20);
            this.cboTaisho1.TabIndex = 12;
            // 
            // colNo
            // 
            this.colNo.DataPropertyName = "No";
            this.colNo.HeaderText = "No";
            this.colNo.Name = "colNo";
            this.colNo.ReadOnly = true;
            this.colNo.Width = 25;
            // 
            // colChk
            // 
            this.colChk.HeaderText = "";
            this.colChk.Name = "colChk";
            this.colChk.Width = 25;
            // 
            // colTaisho
            // 
            this.colTaisho.DataPropertyName = "Taisho";
            this.colTaisho.HeaderText = "対象";
            this.colTaisho.Name = "colTaisho";
            this.colTaisho.ReadOnly = true;
            this.colTaisho.Width = 50;
            // 
            // colSKUCD
            // 
            this.colSKUCD.DataPropertyName = "SKUCD";
            this.colSKUCD.HeaderText = "SKUCD";
            this.colSKUCD.Name = "colSKUCD";
            this.colSKUCD.ReadOnly = true;
            this.colSKUCD.Width = 200;
            // 
            // colJANCD
            // 
            this.colJANCD.DataPropertyName = "JANCD";
            this.colJANCD.HeaderText = "JANCD";
            this.colJANCD.Name = "colJANCD";
            this.colJANCD.ReadOnly = true;
            this.colJANCD.Width = 120;
            // 
            // colMakerCD
            // 
            this.colMakerCD.DataPropertyName = "vc_maker_shohin";
            this.colMakerCD.HeaderText = "メーカー商品CD";
            this.colMakerCD.Name = "colMakerCD";
            this.colMakerCD.ReadOnly = true;
            this.colMakerCD.Width = 120;
            // 
            // colItemName
            // 
            this.colItemName.DataPropertyName = "vm_kihon_shohin";
            this.colItemName.HeaderText = "商品名";
            this.colItemName.Name = "colItemName";
            this.colItemName.ReadOnly = true;
            this.colItemName.Width = 300;
            // 
            // colColorName
            // 
            this.colColorName.DataPropertyName = "vm_color";
            this.colColorName.HeaderText = "カラー";
            this.colColorName.Name = "colColorName";
            this.colColorName.ReadOnly = true;
            // 
            // colSizeName
            // 
            this.colSizeName.DataPropertyName = "vm_size";
            this.colSizeName.HeaderText = "サイズ";
            this.colSizeName.Name = "colSizeName";
            this.colSizeName.ReadOnly = true;
            this.colSizeName.Width = 80;
            // 
            // colRefNo
            // 
            this.colRefNo.DataPropertyName = "RefNo";
            this.colRefNo.DecimalPlace = ((byte)(0));
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            this.colRefNo.DefaultCellStyle = dataGridViewCellStyle3;
            this.colRefNo.HeaderText = "基準数";
            this.colRefNo.MaxInputLength = 32767;
            this.colRefNo.Name = "colRefNo";
            this.colRefNo.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.colRefNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.colRefNo.UseMinus = false;
            this.colRefNo.UseThousandSeparator = true;
            this.colRefNo.Width = 50;
            // 
            // colZaikoKeisan
            // 
            this.colZaikoKeisan.DataPropertyName = "ZaikoKeisan";
            this.colZaikoKeisan.HeaderText = "在庫計算";
            this.colZaikoKeisan.Name = "colZaikoKeisan";
            this.colZaikoKeisan.Width = 150;
            // 
            // colShiiresakiName
            // 
            this.colShiiresakiName.DataPropertyName = "VM_SHIIRESAKI";
            this.colShiiresakiName.HeaderText = "仕入先名";
            this.colShiiresakiName.Name = "colShiiresakiName";
            this.colShiiresakiName.ReadOnly = true;
            // 
            // colBrandName
            // 
            this.colBrandName.DataPropertyName = "vm_brand";
            this.colBrandName.HeaderText = "ブランド名";
            this.colBrandName.Name = "colBrandName";
            this.colBrandName.ReadOnly = true;
            this.colBrandName.Width = 120;
            // 
            // colInstructionNo
            // 
            this.colInstructionNo.DataPropertyName = "InstructionNo";
            this.colInstructionNo.HeaderText = "指示書番号";
            this.colInstructionNo.MaxInputLength = 20;
            this.colInstructionNo.Name = "colInstructionNo";
            // 
            // colRemark
            // 
            this.colRemark.DataPropertyName = "Remark";
            this.colRemark.HeaderText = "備考";
            this.colRemark.MaxInputLength = 600;
            this.colRemark.Name = "colRemark";
            this.colRemark.Width = 300;
            // 
            // colAdminCD
            // 
            this.colAdminCD.DataPropertyName = "AdminCD";
            this.colAdminCD.HeaderText = "AdminCD";
            this.colAdminCD.Name = "colAdminCD";
            this.colAdminCD.Visible = false;
            // 
            // frmPSKS0119M
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1884, 951);
            this.F11_Text = "Excel出力(F11)";
            this.KeyPreview = true;
            this.Location = new System.Drawing.Point(0, 0);
            this.Name = "frmPSKS0119M";
            this.Text = "PSKS0119M";
            this.Load += new System.EventHandler(this.frmPSKS0119M_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmPSKS0119M_KeyDown);
            this.panelBody.ResumeLayout(false);
            this.panelBody.PerformLayout();
            this.panelTitleLeft.ResumeLayout(false);
            this.panelTitle.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.gdvPSKS0119M)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.smS_GridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gdv)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private CustomControls.SMS_Label smS_Label2;
        private CustomControls.SMS_TextBox txtFilePath;
        private CustomControls.SMS_Button btnBrowse;
        private CustomControls.SMS_Button btnImport;
        private CustomControls.SMS_Button btnDisplay;
        private CustomControls.SMS_Label smS_Label3;
        private CustomControls.SMS_ComboBox cboBrand;
        private CustomControls.SMS_ComboBox cboShiiresaki;
        private CustomControls.SMS_ComboBox cboTaisho;
        private CustomControls.SMS_Label smS_Label6;
        private CustomControls.SMS_Label smS_Label5;
        private CustomControls.SMS_Label smS_Label4;
        private CustomControls.SMS_Button btnAllUncheck;
        private CustomControls.SMS_Button btnAllCheck;
        private CustomControls.SMS_Button btnItemSearch;
        private CustomControls.SMS_Button btnSetting;
        private CustomControls.SMS_ComboBox cboZaikoKeisan;
        private CustomControls.SMS_Label smS_Label8;
        private CustomControls.SMS_TextBox txtRefNo;
        private CustomControls.SMS_Label smS_Label7;
        private SMS_GridView gdvPSKS0119M;
        private SMS_GridView smS_GridView1;
        private SMS_GridView gdv;
        private CustomControls.SMS_ComboBox cboTaisho1;
        private CustomControls.SMS_Label smS_Label9;
        private System.Windows.Forms.DataGridViewTextBoxColumn colNo;
        private System.Windows.Forms.DataGridViewCheckBoxColumn colChk;
        private System.Windows.Forms.DataGridViewTextBoxColumn colTaisho;
        private System.Windows.Forms.DataGridViewTextBoxColumn colSKUCD;
        private System.Windows.Forms.DataGridViewTextBoxColumn colJANCD;
        private System.Windows.Forms.DataGridViewTextBoxColumn colMakerCD;
        private System.Windows.Forms.DataGridViewTextBoxColumn colItemName;
        private System.Windows.Forms.DataGridViewTextBoxColumn colColorName;
        private System.Windows.Forms.DataGridViewTextBoxColumn colSizeName;
        private CustomControls.DataGridViewDecimalColumn colRefNo;
        private System.Windows.Forms.DataGridViewComboBoxColumn colZaikoKeisan;
        private System.Windows.Forms.DataGridViewTextBoxColumn colShiiresakiName;
        private System.Windows.Forms.DataGridViewTextBoxColumn colBrandName;
        private System.Windows.Forms.DataGridViewTextBoxColumn colInstructionNo;
        private System.Windows.Forms.DataGridViewTextBoxColumn colRemark;
        private System.Windows.Forms.DataGridViewTextBoxColumn colAdminCD;
    }
}