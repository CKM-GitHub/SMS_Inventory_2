namespace SMS.Inventory
{
    partial class frmPSKS0120S
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
            this.chkMakerLikeSearch = new System.Windows.Forms.CheckBox();
            this.ucCategory = new SMS.CustomControls.UC_Search();
            this.ucKyouGiName = new SMS.CustomControls.UC_Search();
            this.ucSupplier = new SMS.CustomControls.UC_Search();
            this.cboYoyakuFlg = new SMS.CustomControls.SMS_ComboBox();
            this.smS_Label20 = new SMS.CustomControls.SMS_Label();
            this.txtCatalog = new SMS.CustomControls.SMS_TextBox();
            this.smS_Label19 = new SMS.CustomControls.SMS_Label();
            this.smS_Label18 = new SMS.CustomControls.SMS_Label();
            this.smS_Label17 = new SMS.CustomControls.SMS_Label();
            this.smS_Label16 = new SMS.CustomControls.SMS_Label();
            this.cboSpecialFlag = new SMS.CustomControls.SMS_ComboBox();
            this.smS_Label15 = new SMS.CustomControls.SMS_Label();
            this.txtItemName = new SMS.CustomControls.SMS_TextBox();
            this.smS_Label14 = new SMS.CustomControls.SMS_Label();
            this.cboSeason = new SMS.CustomControls.SMS_ComboBox();
            this.cboNendo = new SMS.CustomControls.SMS_ComboBox();
            this.txtJANCD = new SMS.CustomControls.SMS_TextBox();
            this.txtMakerShohinCD = new SMS.CustomControls.SMS_TextBox();
            this.txtItem = new SMS.CustomControls.SMS_TextBox();
            this.btnDisplay = new SMS.CustomControls.SMS_Button();
            this.txtShijisho = new SMS.CustomControls.SMS_TextBox();
            this.ucBrand = new SMS.CustomControls.UC_Search();
            this.smS_Label10 = new SMS.CustomControls.SMS_Label();
            this.smS_Label8 = new SMS.CustomControls.SMS_Label();
            this.smS_Label7 = new SMS.CustomControls.SMS_Label();
            this.smS_Label5 = new SMS.CustomControls.SMS_Label();
            this.smS_Label4 = new SMS.CustomControls.SMS_Label();
            this.smS_Label3 = new SMS.CustomControls.SMS_Label();
            this.smS_Label2 = new SMS.CustomControls.SMS_Label();
            this.dgvPSKS0102S = new SMS.SMS_GridView();
            this.colNo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colOperatorCd = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colAdminCD = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colvm_shiiresaki = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colvm_brand = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colvc_maker_shohin = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colJANCD = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colvm_kihon_shohin = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colvm_color = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colvm_size = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colZaiko_keisan = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colCapitalZaiko = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colToyonakaZaiko = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colKijunsu = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colShijisho = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colMemo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colSKUCD = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.panelBody.SuspendLayout();
            this.panelTitleLeft.SuspendLayout();
            this.panelTitle.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvPSKS0102S)).BeginInit();
            this.SuspendLayout();
            // 
            // panelBody
            // 
            this.panelBody.Controls.Add(this.dgvPSKS0102S);
            this.panelBody.Controls.Add(this.chkMakerLikeSearch);
            this.panelBody.Controls.Add(this.ucCategory);
            this.panelBody.Controls.Add(this.ucKyouGiName);
            this.panelBody.Controls.Add(this.ucSupplier);
            this.panelBody.Controls.Add(this.cboYoyakuFlg);
            this.panelBody.Controls.Add(this.smS_Label20);
            this.panelBody.Controls.Add(this.txtCatalog);
            this.panelBody.Controls.Add(this.smS_Label19);
            this.panelBody.Controls.Add(this.smS_Label18);
            this.panelBody.Controls.Add(this.smS_Label17);
            this.panelBody.Controls.Add(this.smS_Label16);
            this.panelBody.Controls.Add(this.cboSpecialFlag);
            this.panelBody.Controls.Add(this.smS_Label15);
            this.panelBody.Controls.Add(this.txtItemName);
            this.panelBody.Controls.Add(this.smS_Label14);
            this.panelBody.Controls.Add(this.cboSeason);
            this.panelBody.Controls.Add(this.cboNendo);
            this.panelBody.Controls.Add(this.txtJANCD);
            this.panelBody.Controls.Add(this.txtMakerShohinCD);
            this.panelBody.Controls.Add(this.txtItem);
            this.panelBody.Controls.Add(this.btnDisplay);
            this.panelBody.Controls.Add(this.txtShijisho);
            this.panelBody.Controls.Add(this.ucBrand);
            this.panelBody.Controls.Add(this.smS_Label10);
            this.panelBody.Controls.Add(this.smS_Label8);
            this.panelBody.Controls.Add(this.smS_Label7);
            this.panelBody.Controls.Add(this.smS_Label5);
            this.panelBody.Controls.Add(this.smS_Label4);
            this.panelBody.Controls.Add(this.smS_Label3);
            this.panelBody.Controls.Add(this.smS_Label2);
            // 
            // chkMakerLikeSearch
            // 
            this.chkMakerLikeSearch.AutoSize = true;
            this.chkMakerLikeSearch.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkMakerLikeSearch.Location = new System.Drawing.Point(103, 96);
            this.chkMakerLikeSearch.Name = "chkMakerLikeSearch";
            this.chkMakerLikeSearch.Size = new System.Drawing.Size(50, 16);
            this.chkMakerLikeSearch.TabIndex = 114;
            this.chkMakerLikeSearch.Text = "完全";
            this.chkMakerLikeSearch.UseVisualStyleBackColor = true;
            // 
            // ucCategory
            // 
            this.ucCategory.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            this.ucCategory.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ucCategory.Location = new System.Drawing.Point(660, 47);
            this.ucCategory.Margin = new System.Windows.Forms.Padding(0);
            this.ucCategory.Name = "ucCategory";
            this.ucCategory.Size = new System.Drawing.Size(320, 23);
            this.ucCategory.TabIndex = 85;
            this.ucCategory.UC_Code = "";
            this.ucCategory.UC_Code_Width = 60;
            this.ucCategory.UC_Exists = false;
            this.ucCategory.UC_Flag = 0;
            this.ucCategory.UC_IsRequired = false;
            this.ucCategory.UC_Name = "";
            this.ucCategory.UC_Name_Width = 220;
            this.ucCategory.UC_NameVisible = true;
            this.ucCategory.UC_SearchEnable = true;
            this.ucCategory.UC_Type = SMS.CustomControls.UC_Search.Type.Category;
            this.ucCategory.Enter += new System.EventHandler(this.uc_Enter);
            this.ucCategory.Leave += new System.EventHandler(this.uc_Leave);
            // 
            // ucKyouGiName
            // 
            this.ucKyouGiName.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            this.ucKyouGiName.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ucKyouGiName.Location = new System.Drawing.Point(660, 16);
            this.ucKyouGiName.Margin = new System.Windows.Forms.Padding(0);
            this.ucKyouGiName.Name = "ucKyouGiName";
            this.ucKyouGiName.Size = new System.Drawing.Size(320, 23);
            this.ucKyouGiName.TabIndex = 78;
            this.ucKyouGiName.UC_Code = "";
            this.ucKyouGiName.UC_Code_Width = 60;
            this.ucKyouGiName.UC_Exists = false;
            this.ucKyouGiName.UC_Flag = 0;
            this.ucKyouGiName.UC_IsRequired = false;
            this.ucKyouGiName.UC_Name = "";
            this.ucKyouGiName.UC_Name_Width = 220;
            this.ucKyouGiName.UC_NameVisible = true;
            this.ucKyouGiName.UC_SearchEnable = true;
            this.ucKyouGiName.UC_Type = SMS.CustomControls.UC_Search.Type.Sports;
            this.ucKyouGiName.Enter += new System.EventHandler(this.uc_Enter);
            this.ucKyouGiName.Leave += new System.EventHandler(this.uc_Leave);
            // 
            // ucSupplier
            // 
            this.ucSupplier.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            this.ucSupplier.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ucSupplier.Location = new System.Drawing.Point(169, 16);
            this.ucSupplier.Margin = new System.Windows.Forms.Padding(0);
            this.ucSupplier.Name = "ucSupplier";
            this.ucSupplier.Size = new System.Drawing.Size(430, 25);
            this.ucSupplier.TabIndex = 76;
            this.ucSupplier.UC_Code = "";
            this.ucSupplier.UC_Code_Width = 65;
            this.ucSupplier.UC_Exists = false;
            this.ucSupplier.UC_Flag = 0;
            this.ucSupplier.UC_IsRequired = false;
            this.ucSupplier.UC_Name = "";
            this.ucSupplier.UC_Name_Width = 330;
            this.ucSupplier.UC_NameVisible = true;
            this.ucSupplier.UC_SearchEnable = true;
            this.ucSupplier.UC_Type = SMS.CustomControls.UC_Search.Type.Maker;
            this.ucSupplier.Enter += new System.EventHandler(this.uc_Enter);
            this.ucSupplier.Leave += new System.EventHandler(this.uc_Leave);
            // 
            // cboYoyakuFlg
            // 
            this.cboYoyakuFlg.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Append;
            this.cboYoyakuFlg.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cboYoyakuFlg.Cbo_Type = SMS.CustomControls.SMS_ComboBox.cbo_Type.YoyakuFlag;
            this.cboYoyakuFlg.DisplayMember = "vm_nendo";
            this.cboYoyakuFlg.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cboYoyakuFlg.FormattingEnabled = true;
            this.cboYoyakuFlg.Location = new System.Drawing.Point(1541, 21);
            this.cboYoyakuFlg.Name = "cboYoyakuFlg";
            this.cboYoyakuFlg.Size = new System.Drawing.Size(121, 20);
            this.cboYoyakuFlg.TabIndex = 81;
            this.cboYoyakuFlg.ValueMember = "nk_nendo";
            // 
            // smS_Label20
            // 
            this.smS_Label20.AutoSize = true;
            this.smS_Label20.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label20.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label20.Location = new System.Drawing.Point(1465, 24);
            this.smS_Label20.Name = "smS_Label20";
            this.smS_Label20.Size = new System.Drawing.Size(70, 12);
            this.smS_Label20.TabIndex = 113;
            this.smS_Label20.Text = "予約フラグ";
            // 
            // txtCatalog
            // 
            this.txtCatalog.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtCatalog.Decimalplace = ((byte)(0));
            this.txtCatalog.Location = new System.Drawing.Point(1541, 139);
            this.txtCatalog.MaxLength = 30;
            this.txtCatalog.Name = "txtCatalog";
            this.txtCatalog.Size = new System.Drawing.Size(250, 19);
            this.txtCatalog.TabIndex = 98;
            this.txtCatalog.txt_Type = SMS.CustomControls.SMS_TextBox.Type.Hiragana;
            this.txtCatalog.UseMinus = false;
            this.txtCatalog.UseThousandSeparator = true;
            // 
            // smS_Label19
            // 
            this.smS_Label19.AutoSize = true;
            this.smS_Label19.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label19.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label19.Location = new System.Drawing.Point(1451, 141);
            this.smS_Label19.Name = "smS_Label19";
            this.smS_Label19.Size = new System.Drawing.Size(83, 12);
            this.smS_Label19.TabIndex = 112;
            this.smS_Label19.Text = "カタログ情報";
            // 
            // smS_Label18
            // 
            this.smS_Label18.AutoSize = true;
            this.smS_Label18.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label18.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label18.Location = new System.Drawing.Point(119, 21);
            this.smS_Label18.Name = "smS_Label18";
            this.smS_Label18.Size = new System.Drawing.Size(44, 12);
            this.smS_Label18.TabIndex = 111;
            this.smS_Label18.Text = "仕入先";
            // 
            // smS_Label17
            // 
            this.smS_Label17.AutoSize = true;
            this.smS_Label17.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label17.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label17.Location = new System.Drawing.Point(625, 52);
            this.smS_Label17.Name = "smS_Label17";
            this.smS_Label17.Size = new System.Drawing.Size(31, 12);
            this.smS_Label17.TabIndex = 110;
            this.smS_Label17.Text = "分類";
            // 
            // smS_Label16
            // 
            this.smS_Label16.AutoSize = true;
            this.smS_Label16.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label16.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label16.Location = new System.Drawing.Point(625, 21);
            this.smS_Label16.Name = "smS_Label16";
            this.smS_Label16.Size = new System.Drawing.Size(31, 12);
            this.smS_Label16.TabIndex = 109;
            this.smS_Label16.Text = "競技";
            // 
            // cboSpecialFlag
            // 
            this.cboSpecialFlag.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Append;
            this.cboSpecialFlag.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cboSpecialFlag.Cbo_Type = SMS.CustomControls.SMS_ComboBox.cbo_Type.SpecialFlag;
            this.cboSpecialFlag.DisplayMember = "vm_nendo";
            this.cboSpecialFlag.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cboSpecialFlag.FormattingEnabled = true;
            this.cboSpecialFlag.Location = new System.Drawing.Point(1541, 51);
            this.cboSpecialFlag.Name = "cboSpecialFlag";
            this.cboSpecialFlag.Size = new System.Drawing.Size(121, 20);
            this.cboSpecialFlag.TabIndex = 90;
            this.cboSpecialFlag.ValueMember = "nk_nendo";
            // 
            // smS_Label15
            // 
            this.smS_Label15.AutoSize = true;
            this.smS_Label15.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label15.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label15.Location = new System.Drawing.Point(1464, 56);
            this.smS_Label15.Name = "smS_Label15";
            this.smS_Label15.Size = new System.Drawing.Size(70, 12);
            this.smS_Label15.TabIndex = 108;
            this.smS_Label15.Text = "特記フラグ";
            // 
            // txtItemName
            // 
            this.txtItemName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtItemName.Decimalplace = ((byte)(0));
            this.txtItemName.Location = new System.Drawing.Point(1086, 19);
            this.txtItemName.MaxLength = 128;
            this.txtItemName.Name = "txtItemName";
            this.txtItemName.Size = new System.Drawing.Size(230, 19);
            this.txtItemName.TabIndex = 80;
            this.txtItemName.txt_Type = SMS.CustomControls.SMS_TextBox.Type.Hiragana;
            this.txtItemName.UseMinus = false;
            this.txtItemName.UseThousandSeparator = true;
            // 
            // smS_Label14
            // 
            this.smS_Label14.AutoSize = true;
            this.smS_Label14.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label14.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label14.Location = new System.Drawing.Point(1035, 21);
            this.smS_Label14.Name = "smS_Label14";
            this.smS_Label14.Size = new System.Drawing.Size(44, 12);
            this.smS_Label14.TabIndex = 107;
            this.smS_Label14.Text = "商品名";
            // 
            // cboSeason
            // 
            this.cboSeason.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Append;
            this.cboSeason.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cboSeason.Cbo_Type = SMS.CustomControls.SMS_ComboBox.cbo_Type.m_season;
            this.cboSeason.DisplayMember = "vm_nendo";
            this.cboSeason.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cboSeason.FormattingEnabled = true;
            this.cboSeason.Location = new System.Drawing.Point(1541, 108);
            this.cboSeason.Name = "cboSeason";
            this.cboSeason.Size = new System.Drawing.Size(121, 20);
            this.cboSeason.TabIndex = 97;
            this.cboSeason.ValueMember = "nk_nendo";
            // 
            // cboNendo
            // 
            this.cboNendo.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Append;
            this.cboNendo.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cboNendo.Cbo_Type = SMS.CustomControls.SMS_ComboBox.cbo_Type.m_nendo;
            this.cboNendo.DisplayMember = "vm_nendo";
            this.cboNendo.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cboNendo.FormattingEnabled = true;
            this.cboNendo.Location = new System.Drawing.Point(1541, 81);
            this.cboNendo.Name = "cboNendo";
            this.cboNendo.Size = new System.Drawing.Size(121, 20);
            this.cboNendo.TabIndex = 96;
            this.cboNendo.ValueMember = "nk_nendo";
            // 
            // txtJANCD
            // 
            this.txtJANCD.Decimalplace = ((byte)(0));
            this.txtJANCD.Location = new System.Drawing.Point(1086, 79);
            this.txtJANCD.Multiline = true;
            this.txtJANCD.Name = "txtJANCD";
            this.txtJANCD.Size = new System.Drawing.Size(350, 80);
            this.txtJANCD.TabIndex = 95;
            this.txtJANCD.txt_Type = SMS.CustomControls.SMS_TextBox.Type.Normal;
            this.txtJANCD.UseMinus = false;
            this.txtJANCD.UseThousandSeparator = true;
            // 
            // txtMakerShohinCD
            // 
            this.txtMakerShohinCD.Decimalplace = ((byte)(0));
            this.txtMakerShohinCD.Location = new System.Drawing.Point(169, 76);
            this.txtMakerShohinCD.Multiline = true;
            this.txtMakerShohinCD.Name = "txtMakerShohinCD";
            this.txtMakerShohinCD.Size = new System.Drawing.Size(350, 80);
            this.txtMakerShohinCD.TabIndex = 92;
            this.txtMakerShohinCD.txt_Type = SMS.CustomControls.SMS_TextBox.Type.Normal;
            this.txtMakerShohinCD.UseMinus = false;
            this.txtMakerShohinCD.UseThousandSeparator = true;
            // 
            // txtItem
            // 
            this.txtItem.Decimalplace = ((byte)(0));
            this.txtItem.Location = new System.Drawing.Point(660, 78);
            this.txtItem.Multiline = true;
            this.txtItem.Name = "txtItem";
            this.txtItem.Size = new System.Drawing.Size(350, 80);
            this.txtItem.TabIndex = 94;
            this.txtItem.txt_Type = SMS.CustomControls.SMS_TextBox.Type.Normal;
            this.txtItem.UseMinus = false;
            this.txtItem.UseThousandSeparator = true;
            // 
            // btnDisplay
            // 
            this.btnDisplay.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.btnDisplay.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnDisplay.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnDisplay.Location = new System.Drawing.Point(1707, 191);
            this.btnDisplay.Name = "btnDisplay";
            this.btnDisplay.Size = new System.Drawing.Size(109, 28);
            this.btnDisplay.TabIndex = 99;
            this.btnDisplay.Text = "表示（F10）";
            this.btnDisplay.UseVisualStyleBackColor = false;
            this.btnDisplay.Click += new System.EventHandler(this.btnDisplay_Click);
            // 
            // txtShijisho
            // 
            this.txtShijisho.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtShijisho.Decimalplace = ((byte)(0));
            this.txtShijisho.Location = new System.Drawing.Point(1086, 50);
            this.txtShijisho.MaxLength = 30;
            this.txtShijisho.Name = "txtShijisho";
            this.txtShijisho.Size = new System.Drawing.Size(250, 19);
            this.txtShijisho.TabIndex = 87;
            this.txtShijisho.txt_Type = SMS.CustomControls.SMS_TextBox.Type.Hiragana;
            this.txtShijisho.UseMinus = false;
            this.txtShijisho.UseThousandSeparator = true;
            // 
            // ucBrand
            // 
            this.ucBrand.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            this.ucBrand.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ucBrand.Location = new System.Drawing.Point(169, 47);
            this.ucBrand.Margin = new System.Windows.Forms.Padding(0);
            this.ucBrand.Name = "ucBrand";
            this.ucBrand.Size = new System.Drawing.Size(330, 21);
            this.ucBrand.TabIndex = 84;
            this.ucBrand.UC_Code = "";
            this.ucBrand.UC_Code_Width = 40;
            this.ucBrand.UC_Exists = false;
            this.ucBrand.UC_Flag = 0;
            this.ucBrand.UC_IsRequired = false;
            this.ucBrand.UC_Name = "";
            this.ucBrand.UC_Name_Width = 250;
            this.ucBrand.UC_NameVisible = true;
            this.ucBrand.UC_SearchEnable = true;
            this.ucBrand.UC_Type = SMS.CustomControls.UC_Search.Type.Brand;
            this.ucBrand.Enter += new System.EventHandler(this.uc_Enter);
            this.ucBrand.Leave += new System.EventHandler(this.uc_Leave);
            // 
            // smS_Label10
            // 
            this.smS_Label10.AutoSize = true;
            this.smS_Label10.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label10.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label10.Location = new System.Drawing.Point(1478, 111);
            this.smS_Label10.Name = "smS_Label10";
            this.smS_Label10.Size = new System.Drawing.Size(57, 12);
            this.smS_Label10.TabIndex = 93;
            this.smS_Label10.Text = "シーズン";
            // 
            // smS_Label8
            // 
            this.smS_Label8.AutoSize = true;
            this.smS_Label8.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label8.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label8.Location = new System.Drawing.Point(1503, 84);
            this.smS_Label8.Name = "smS_Label8";
            this.smS_Label8.Size = new System.Drawing.Size(31, 12);
            this.smS_Label8.TabIndex = 88;
            this.smS_Label8.Text = "年度";
            // 
            // smS_Label7
            // 
            this.smS_Label7.AutoSize = true;
            this.smS_Label7.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label7.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label7.Location = new System.Drawing.Point(1009, 52);
            this.smS_Label7.Name = "smS_Label7";
            this.smS_Label7.Size = new System.Drawing.Size(70, 12);
            this.smS_Label7.TabIndex = 86;
            this.smS_Label7.Text = "指示書番号";
            // 
            // smS_Label5
            // 
            this.smS_Label5.AutoSize = true;
            this.smS_Label5.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label5.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label5.Location = new System.Drawing.Point(1040, 81);
            this.smS_Label5.Name = "smS_Label5";
            this.smS_Label5.Size = new System.Drawing.Size(40, 12);
            this.smS_Label5.TabIndex = 82;
            this.smS_Label5.Text = "JANCD";
            // 
            // smS_Label4
            // 
            this.smS_Label4.AutoSize = true;
            this.smS_Label4.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label4.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label4.Location = new System.Drawing.Point(66, 81);
            this.smS_Label4.Name = "smS_Label4";
            this.smS_Label4.Size = new System.Drawing.Size(97, 12);
            this.smS_Label4.TabIndex = 79;
            this.smS_Label4.Text = "メーカー商品CD";
            // 
            // smS_Label3
            // 
            this.smS_Label3.AutoSize = true;
            this.smS_Label3.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label3.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label3.Location = new System.Drawing.Point(613, 79);
            this.smS_Label3.Name = "smS_Label3";
            this.smS_Label3.Size = new System.Drawing.Size(33, 12);
            this.smS_Label3.TabIndex = 77;
            this.smS_Label3.Text = "ITEM";
            // 
            // smS_Label2
            // 
            this.smS_Label2.AutoSize = true;
            this.smS_Label2.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label2.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label2.Location = new System.Drawing.Point(105, 52);
            this.smS_Label2.Name = "smS_Label2";
            this.smS_Label2.Size = new System.Drawing.Size(57, 12);
            this.smS_Label2.TabIndex = 75;
            this.smS_Label2.Text = "ブランド";
            // 
            // dgvPSKS0102S
            // 
            this.dgvPSKS0102S.AllowUserToAddRows = false;
            this.dgvPSKS0102S.AllowUserToDeleteRows = false;
            this.dgvPSKS0102S.AllowUserToResizeRows = false;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(221)))), ((int)(((byte)(235)))), ((int)(((byte)(247)))));
            this.dgvPSKS0102S.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dgvPSKS0102S.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            dataGridViewCellStyle2.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgvPSKS0102S.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle2;
            this.dgvPSKS0102S.ColumnHeadersHeight = 25;
            this.dgvPSKS0102S.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colNo,
            this.colOperatorCd,
            this.colAdminCD,
            this.colvm_shiiresaki,
            this.colvm_brand,
            this.colvc_maker_shohin,
            this.colJANCD,
            this.colvm_kihon_shohin,
            this.colvm_color,
            this.colvm_size,
            this.colZaiko_keisan,
            this.colCapitalZaiko,
            this.colToyonakaZaiko,
            this.colKijunsu,
            this.colShijisho,
            this.colMemo,
            this.colSKUCD});
            dataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle6.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle6.Font = new System.Drawing.Font("MS Gothic", 9F);
            dataGridViewCellStyle6.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgvPSKS0102S.DefaultCellStyle = dataGridViewCellStyle6;
            this.dgvPSKS0102S.EnableHeadersVisualStyles = false;
            this.dgvPSKS0102S.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            this.dgvPSKS0102S.IsCellValid = false;
            this.dgvPSKS0102S.Location = new System.Drawing.Point(46, 238);
            this.dgvPSKS0102S.Name = "dgvPSKS0102S";
            this.dgvPSKS0102S.RowHeadersVisible = false;
            this.dgvPSKS0102S.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            this.dgvPSKS0102S.Size = new System.Drawing.Size(1770, 640);
            this.dgvPSKS0102S.StandardTab = true;
            this.dgvPSKS0102S.TabIndex = 115;
            this.dgvPSKS0102S.CellEnter += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvPSKS0102S_CellEnter);
            this.dgvPSKS0102S.CellFormatting += new System.Windows.Forms.DataGridViewCellFormattingEventHandler(this.dgvPSKS0102S_CellFormatting);
            // 
            // colNo
            // 
            this.colNo.DataPropertyName = "No";
            this.colNo.HeaderText = "No";
            this.colNo.Name = "colNo";
            this.colNo.ReadOnly = true;
            this.colNo.Width = 50;
            // 
            // colOperatorCd
            // 
            this.colOperatorCd.HeaderText = "OperatorCD";
            this.colOperatorCd.Name = "colOperatorCd";
            this.colOperatorCd.ReadOnly = true;
            this.colOperatorCd.Visible = false;
            // 
            // colAdminCD
            // 
            this.colAdminCD.HeaderText = "AdminCD";
            this.colAdminCD.Name = "colAdminCD";
            this.colAdminCD.ReadOnly = true;
            this.colAdminCD.Visible = false;
            // 
            // colvm_shiiresaki
            // 
            this.colvm_shiiresaki.DataPropertyName = "vm_shiiresaki";
            this.colvm_shiiresaki.HeaderText = "仕入先名";
            this.colvm_shiiresaki.Name = "colvm_shiiresaki";
            this.colvm_shiiresaki.ReadOnly = true;
            // 
            // colvm_brand
            // 
            this.colvm_brand.DataPropertyName = "vm_brand";
            this.colvm_brand.HeaderText = "ブランド名";
            this.colvm_brand.Name = "colvm_brand";
            this.colvm_brand.ReadOnly = true;
            // 
            // colvc_maker_shohin
            // 
            this.colvc_maker_shohin.DataPropertyName = "vc_maker_shohin";
            this.colvc_maker_shohin.HeaderText = "メーカー商品CD";
            this.colvc_maker_shohin.Name = "colvc_maker_shohin";
            this.colvc_maker_shohin.ReadOnly = true;
            this.colvc_maker_shohin.Width = 110;
            // 
            // colJANCD
            // 
            this.colJANCD.DataPropertyName = "JANCD";
            this.colJANCD.HeaderText = "JANCD";
            this.colJANCD.Name = "colJANCD";
            this.colJANCD.ReadOnly = true;
            // 
            // colvm_kihon_shohin
            // 
            this.colvm_kihon_shohin.DataPropertyName = "vm_kihon_shohin";
            this.colvm_kihon_shohin.HeaderText = "商品名";
            this.colvm_kihon_shohin.Name = "colvm_kihon_shohin";
            this.colvm_kihon_shohin.ReadOnly = true;
            this.colvm_kihon_shohin.Width = 200;
            // 
            // colvm_color
            // 
            this.colvm_color.DataPropertyName = "vm_color";
            this.colvm_color.HeaderText = "カラー";
            this.colvm_color.Name = "colvm_color";
            this.colvm_color.ReadOnly = true;
            this.colvm_color.Width = 150;
            // 
            // colvm_size
            // 
            this.colvm_size.DataPropertyName = "vm_size";
            this.colvm_size.HeaderText = "サイズ";
            this.colvm_size.Name = "colvm_size";
            this.colvm_size.ReadOnly = true;
            this.colvm_size.Width = 150;
            // 
            // colZaiko_keisan
            // 
            this.colZaiko_keisan.DataPropertyName = "Zaiko_keisan";
            this.colZaiko_keisan.HeaderText = "Zaiko_keisan";
            this.colZaiko_keisan.Name = "colZaiko_keisan";
            this.colZaiko_keisan.ReadOnly = true;
            this.colZaiko_keisan.Visible = false;
            // 
            // colCapitalZaiko
            // 
            this.colCapitalZaiko.DataPropertyName = "CapitalZaiko";
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            this.colCapitalZaiko.DefaultCellStyle = dataGridViewCellStyle3;
            this.colCapitalZaiko.HeaderText = "Web在庫";
            this.colCapitalZaiko.Name = "colCapitalZaiko";
            this.colCapitalZaiko.ReadOnly = true;
            this.colCapitalZaiko.Width = 60;
            // 
            // colToyonakaZaiko
            // 
            this.colToyonakaZaiko.DataPropertyName = "ToyonakaZaiko";
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            this.colToyonakaZaiko.DefaultCellStyle = dataGridViewCellStyle4;
            this.colToyonakaZaiko.HeaderText = "豊中";
            this.colToyonakaZaiko.Name = "colToyonakaZaiko";
            this.colToyonakaZaiko.ReadOnly = true;
            this.colToyonakaZaiko.Width = 60;
            // 
            // colKijunsu
            // 
            this.colKijunsu.DataPropertyName = "Kijunsu";
            dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            this.colKijunsu.DefaultCellStyle = dataGridViewCellStyle5;
            this.colKijunsu.HeaderText = "基準数";
            this.colKijunsu.Name = "colKijunsu";
            this.colKijunsu.ReadOnly = true;
            this.colKijunsu.Width = 60;
            // 
            // colShijisho
            // 
            this.colShijisho.DataPropertyName = "Shijisho";
            this.colShijisho.HeaderText = "指示書番号";
            this.colShijisho.MaxInputLength = 20;
            this.colShijisho.Name = "colShijisho";
            this.colShijisho.Width = 150;
            // 
            // colMemo
            // 
            this.colMemo.DataPropertyName = "Memo";
            this.colMemo.HeaderText = "備考";
            this.colMemo.MaxInputLength = 600;
            this.colMemo.Name = "colMemo";
            this.colMemo.Width = 300;
            // 
            // colSKUCD
            // 
            this.colSKUCD.DataPropertyName = "SKUCD";
            this.colSKUCD.HeaderText = "SKUCD";
            this.colSKUCD.Name = "colSKUCD";
            this.colSKUCD.ReadOnly = true;
            this.colSKUCD.Width = 160;
            // 
            // frmPSKS0120S
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1884, 951);
            this.KeyPreview = true;
            this.Location = new System.Drawing.Point(0, 0);
            this.Name = "frmPSKS0120S";
            this.Text = "PSKS0120S";
            this.Load += new System.EventHandler(this.frmPSKS0120S_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmPSKS0120S_KeyDown);
            this.panelBody.ResumeLayout(false);
            this.panelBody.PerformLayout();
            this.panelTitleLeft.ResumeLayout(false);
            this.panelTitle.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvPSKS0102S)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.CheckBox chkMakerLikeSearch;
        private CustomControls.UC_Search ucCategory;
        private CustomControls.UC_Search ucKyouGiName;
        private CustomControls.UC_Search ucSupplier;
        private CustomControls.SMS_ComboBox cboYoyakuFlg;
        private CustomControls.SMS_Label smS_Label20;
        private CustomControls.SMS_TextBox txtCatalog;
        private CustomControls.SMS_Label smS_Label19;
        private CustomControls.SMS_Label smS_Label18;
        private CustomControls.SMS_Label smS_Label17;
        private CustomControls.SMS_Label smS_Label16;
        private CustomControls.SMS_ComboBox cboSpecialFlag;
        private CustomControls.SMS_Label smS_Label15;
        private CustomControls.SMS_TextBox txtItemName;
        private CustomControls.SMS_Label smS_Label14;
        private CustomControls.SMS_ComboBox cboSeason;
        private CustomControls.SMS_ComboBox cboNendo;
        private CustomControls.SMS_TextBox txtJANCD;
        private CustomControls.SMS_TextBox txtMakerShohinCD;
        private CustomControls.SMS_TextBox txtItem;
        private CustomControls.SMS_Button btnDisplay;
        private CustomControls.SMS_TextBox txtShijisho;
        private CustomControls.UC_Search ucBrand;
        private CustomControls.SMS_Label smS_Label10;
        private CustomControls.SMS_Label smS_Label8;
        private CustomControls.SMS_Label smS_Label7;
        private CustomControls.SMS_Label smS_Label5;
        private CustomControls.SMS_Label smS_Label4;
        private CustomControls.SMS_Label smS_Label3;
        private CustomControls.SMS_Label smS_Label2;
        private SMS_GridView dgvPSKS0102S;
        private System.Windows.Forms.DataGridViewTextBoxColumn colNo;
        private System.Windows.Forms.DataGridViewTextBoxColumn colOperatorCd;
        private System.Windows.Forms.DataGridViewTextBoxColumn colAdminCD;
        private System.Windows.Forms.DataGridViewTextBoxColumn colvm_shiiresaki;
        private System.Windows.Forms.DataGridViewTextBoxColumn colvm_brand;
        private System.Windows.Forms.DataGridViewTextBoxColumn colvc_maker_shohin;
        private System.Windows.Forms.DataGridViewTextBoxColumn colJANCD;
        private System.Windows.Forms.DataGridViewTextBoxColumn colvm_kihon_shohin;
        private System.Windows.Forms.DataGridViewTextBoxColumn colvm_color;
        private System.Windows.Forms.DataGridViewTextBoxColumn colvm_size;
        private System.Windows.Forms.DataGridViewTextBoxColumn colZaiko_keisan;
        private System.Windows.Forms.DataGridViewTextBoxColumn colCapitalZaiko;
        private System.Windows.Forms.DataGridViewTextBoxColumn colToyonakaZaiko;
        private System.Windows.Forms.DataGridViewTextBoxColumn colKijunsu;
        private System.Windows.Forms.DataGridViewTextBoxColumn colShijisho;
        private System.Windows.Forms.DataGridViewTextBoxColumn colMemo;
        private System.Windows.Forms.DataGridViewTextBoxColumn colSKUCD;

    }
}