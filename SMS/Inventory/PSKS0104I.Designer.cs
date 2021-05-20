namespace SMS.Inventory
{
    partial class frmPSKS0104I
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle7 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
            this.btnAdd = new SMS.CustomControls.SMS_Button();
            this.smS_Label15 = new SMS.CustomControls.SMS_Label();
            this.ucJANCD2 = new SMS.CustomControls.UC_Search();
            this.smS_Label14 = new SMS.CustomControls.SMS_Label();
            this.btnMiKakutei = new SMS.CustomControls.SMS_Button();
            this.btnKakutei = new SMS.CustomControls.SMS_Button();
            this.btnDisplay = new SMS.CustomControls.SMS_Button();
            this.txtCatalogue = new SMS.CustomControls.SMS_TextBox();
            this.smS_Label13 = new SMS.CustomControls.SMS_Label();
            this.smS_Label12 = new SMS.CustomControls.SMS_Label();
            this.ucSKUCD = new SMS.CustomControls.UC_Search();
            this.chkZaikoSuZero = new System.Windows.Forms.CheckBox();
            this.chkJANCDcut = new System.Windows.Forms.CheckBox();
            this.smS_Label11 = new SMS.CustomControls.SMS_Label();
            this.smS_Label10 = new SMS.CustomControls.SMS_Label();
            this.lbl1 = new SMS.CustomControls.SMS_Label();
            this.smS_Label8 = new SMS.CustomControls.SMS_Label();
            this.rdoMakerItemCD = new System.Windows.Forms.RadioButton();
            this.rdoItem = new System.Windows.Forms.RadioButton();
            this.chkSearch = new System.Windows.Forms.CheckBox();
            this.ucJANCD1 = new SMS.CustomControls.UC_Search();
            this.ucMakerItem = new SMS.CustomControls.UC_Search();
            this.ucITEM = new SMS.CustomControls.UC_Search();
            this.smS_Label7 = new SMS.CustomControls.SMS_Label();
            this.smS_Label6 = new SMS.CustomControls.SMS_Label();
            this.smS_Label5 = new SMS.CustomControls.SMS_Label();
            this.smS_Label4 = new SMS.CustomControls.SMS_Label();
            this.smS_Label3 = new SMS.CustomControls.SMS_Label();
            this.smS_Label2 = new SMS.CustomControls.SMS_Label();
            this.smS_Label16 = new SMS.CustomControls.SMS_Label();
            this.dgvPSK0104I = new SMS.SMS_GridView();
            this.colNo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colFixedDateTime = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colJANCD = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colItemName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colToriKomiDateTime = new SMS.CustomControls.DataGridViewMaskedTextColumn();
            this.colToriKomiNo = new SMS.CustomControls.DataGridViewDecimalColumn();
            this.colKakuteiZaikoSu = new SMS.CustomControls.DataGridViewDecimalColumn();
            this.colNouki = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colTokukiJikou = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colAdminCD = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.coluchiwake = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.smS_Label17 = new SMS.CustomControls.SMS_Label();
            this.smS_Label18 = new SMS.CustomControls.SMS_Label();
            this.lblKakuteiMode = new System.Windows.Forms.Label();
            this.smS_Label19 = new SMS.CustomControls.SMS_Label();
            this.txtZaikoSu = new SMS.CustomControls.SMS_TextBox();
            this.txtInportDate = new CKM_Controls.CKM_TextBox();
            this.cboNendo = new SMS.CustomControls.SMS_ComboBox();
            this.cboSeason = new SMS.CustomControls.SMS_ComboBox();
            this.ucSupplier = new SMS.CustomControls.UC_Search();
            this.ucBrand = new SMS.CustomControls.UC_Search();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panelBody.SuspendLayout();
            this.panelTitleLeft.SuspendLayout();
            this.panelTitle.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvPSK0104I)).BeginInit();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // panelBody
            // 
            this.panelBody.Controls.Add(this.panel1);
            this.panelBody.Controls.Add(this.txtZaikoSu);
            this.panelBody.Controls.Add(this.smS_Label19);
            this.panelBody.Controls.Add(this.lblKakuteiMode);
            this.panelBody.Controls.Add(this.smS_Label18);
            this.panelBody.Controls.Add(this.smS_Label16);
            this.panelBody.Controls.Add(this.dgvPSK0104I);
            this.panelBody.Controls.Add(this.btnAdd);
            this.panelBody.Controls.Add(this.smS_Label15);
            this.panelBody.Controls.Add(this.ucJANCD2);
            this.panelBody.Controls.Add(this.smS_Label14);
            this.panelBody.Controls.Add(this.btnMiKakutei);
            this.panelBody.Controls.Add(this.btnKakutei);
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
            // btnAdd
            // 
            this.btnAdd.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.btnAdd.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnAdd.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnAdd.Location = new System.Drawing.Point(1238, 226);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(100, 30);
            this.btnAdd.TabIndex = 15;
            this.btnAdd.Text = "追加(F10)";
            this.btnAdd.UseVisualStyleBackColor = false;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // smS_Label15
            // 
            this.smS_Label15.AutoSize = true;
            this.smS_Label15.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label15.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label15.Location = new System.Drawing.Point(620, 226);
            this.smS_Label15.Name = "smS_Label15";
            this.smS_Label15.Size = new System.Drawing.Size(44, 12);
            this.smS_Label15.TabIndex = 69;
            this.smS_Label15.Text = "在庫数";
            // 
            // ucJANCD2
            // 
            this.ucJANCD2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            this.ucJANCD2.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ucJANCD2.Location = new System.Drawing.Point(667, 196);
            this.ucJANCD2.Margin = new System.Windows.Forms.Padding(0);
            this.ucJANCD2.Name = "ucJANCD2";
            this.ucJANCD2.Size = new System.Drawing.Size(710, 23);
            this.ucJANCD2.TabIndex = 13;
            this.ucJANCD2.UC_Code = "";
            this.ucJANCD2.UC_Code_Width = 90;
            this.ucJANCD2.UC_Exists = false;
            this.ucJANCD2.UC_Flag = 3;
            this.ucJANCD2.UC_IsRequired = false;
            this.ucJANCD2.UC_Name = "";
            this.ucJANCD2.UC_Name_Width = 549;
            this.ucJANCD2.UC_NameVisible = true;
            this.ucJANCD2.UC_SearchEnable = true;
            this.ucJANCD2.UC_Type = SMS.CustomControls.UC_Search.Type.JANCD;
            this.ucJANCD2.Enter += new System.EventHandler(this.uc_Enter);
            this.ucJANCD2.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ucJANCD2_KeyDown);
            this.ucJANCD2.Leave += new System.EventHandler(this.uc_Leave);
            // 
            // smS_Label14
            // 
            this.smS_Label14.AutoSize = true;
            this.smS_Label14.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label14.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label14.Location = new System.Drawing.Point(624, 202);
            this.smS_Label14.Name = "smS_Label14";
            this.smS_Label14.Size = new System.Drawing.Size(40, 12);
            this.smS_Label14.TabIndex = 67;
            this.smS_Label14.Text = "JANCD";
            // 
            // btnMiKakutei
            // 
            this.btnMiKakutei.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.btnMiKakutei.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnMiKakutei.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnMiKakutei.Location = new System.Drawing.Point(375, 219);
            this.btnMiKakutei.Name = "btnMiKakutei";
            this.btnMiKakutei.Size = new System.Drawing.Size(100, 30);
            this.btnMiKakutei.TabIndex = 15;
            this.btnMiKakutei.Text = "未確定";
            this.btnMiKakutei.UseVisualStyleBackColor = false;
            this.btnMiKakutei.Click += new System.EventHandler(this.btnKakuteiMode);
            // 
            // btnKakutei
            // 
            this.btnKakutei.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.btnKakutei.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnKakutei.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnKakutei.Location = new System.Drawing.Point(275, 219);
            this.btnKakutei.Name = "btnKakutei";
            this.btnKakutei.Size = new System.Drawing.Size(100, 30);
            this.btnKakutei.TabIndex = 14;
            this.btnKakutei.Text = "確定";
            this.btnKakutei.UseVisualStyleBackColor = false;
            this.btnKakutei.Click += new System.EventHandler(this.btnKakuteiMode);
            // 
            // btnDisplay
            // 
            this.btnDisplay.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.btnDisplay.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnDisplay.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnDisplay.Location = new System.Drawing.Point(1231, 129);
            this.btnDisplay.Name = "btnDisplay";
            this.btnDisplay.Size = new System.Drawing.Size(100, 30);
            this.btnDisplay.TabIndex = 10;
            this.btnDisplay.Text = "表示(F11)";
            this.btnDisplay.UseVisualStyleBackColor = false;
            this.btnDisplay.Click += new System.EventHandler(this.btnDisplay_Click);
            // 
            // txtCatalogue
            // 
            this.txtCatalogue.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtCatalogue.Decimalplace = ((byte)(0));
            this.txtCatalogue.Location = new System.Drawing.Point(668, 34);
            this.txtCatalogue.MaxLength = 30;
            this.txtCatalogue.Name = "txtCatalogue";
            this.txtCatalogue.Size = new System.Drawing.Size(220, 19);
            this.txtCatalogue.TabIndex = 3;
            this.txtCatalogue.txt_Type = SMS.CustomControls.SMS_TextBox.Type.Normal;
            this.txtCatalogue.UseMinus = false;
            this.txtCatalogue.UseThousandSeparator = true;
            // 
            // smS_Label13
            // 
            this.smS_Label13.AutoSize = true;
            this.smS_Label13.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label13.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label13.Location = new System.Drawing.Point(1103, 37);
            this.smS_Label13.Name = "smS_Label13";
            this.smS_Label13.Size = new System.Drawing.Size(57, 12);
            this.smS_Label13.TabIndex = 60;
            this.smS_Label13.Text = "シーズン";
            // 
            // smS_Label12
            // 
            this.smS_Label12.AutoSize = true;
            this.smS_Label12.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label12.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label12.Location = new System.Drawing.Point(921, 37);
            this.smS_Label12.Name = "smS_Label12";
            this.smS_Label12.Size = new System.Drawing.Size(31, 12);
            this.smS_Label12.TabIndex = 58;
            this.smS_Label12.Text = "年度";
            // 
            // ucSKUCD
            // 
            this.ucSKUCD.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            this.ucSKUCD.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ucSKUCD.Location = new System.Drawing.Point(668, 57);
            this.ucSKUCD.Margin = new System.Windows.Forms.Padding(0);
            this.ucSKUCD.Name = "ucSKUCD";
            this.ucSKUCD.Size = new System.Drawing.Size(255, 23);
            this.ucSKUCD.TabIndex = 7;
            this.ucSKUCD.UC_Code = "";
            this.ucSKUCD.UC_Code_Width = 220;
            this.ucSKUCD.UC_Exists = false;
            this.ucSKUCD.UC_Flag = 4;
            this.ucSKUCD.UC_IsRequired = false;
            this.ucSKUCD.UC_Name = "";
            this.ucSKUCD.UC_Name_Width = 281;
            this.ucSKUCD.UC_NameVisible = false;
            this.ucSKUCD.UC_SearchEnable = true;
            this.ucSKUCD.UC_Type = SMS.CustomControls.UC_Search.Type.SKUCD;
            this.ucSKUCD.Enter += new System.EventHandler(this.uc_Enter);
            this.ucSKUCD.Leave += new System.EventHandler(this.uc_Leave);
            // 
            // chkZaikoSuZero
            // 
            this.chkZaikoSuZero.AutoSize = true;
            this.chkZaikoSuZero.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkZaikoSuZero.Location = new System.Drawing.Point(1190, 8);
            this.chkZaikoSuZero.Name = "chkZaikoSuZero";
            this.chkZaikoSuZero.Size = new System.Drawing.Size(141, 16);
            this.chkZaikoSuZero.TabIndex = 56;
            this.chkZaikoSuZero.Text = "在庫数が０になった";
            this.chkZaikoSuZero.UseVisualStyleBackColor = true;
            // 
            // chkJANCDcut
            // 
            this.chkJANCDcut.AutoSize = true;
            this.chkJANCDcut.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkJANCDcut.Location = new System.Drawing.Point(1060, 8);
            this.chkJANCDcut.Name = "chkJANCDcut";
            this.chkJANCDcut.Size = new System.Drawing.Size(124, 16);
            this.chkJANCDcut.TabIndex = 55;
            this.chkJANCDcut.Text = "仕入先で初JANCD";
            this.chkJANCDcut.UseVisualStyleBackColor = true;
            // 
            // smS_Label11
            // 
            this.smS_Label11.AutoSize = true;
            this.smS_Label11.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label11.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label11.Location = new System.Drawing.Point(997, 9);
            this.smS_Label11.Name = "smS_Label11";
            this.smS_Label11.Size = new System.Drawing.Size(57, 12);
            this.smS_Label11.TabIndex = 54;
            this.smS_Label11.Text = "表示対象";
            // 
            // smS_Label10
            // 
            this.smS_Label10.AutoSize = true;
            this.smS_Label10.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label10.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label10.Location = new System.Drawing.Point(622, 63);
            this.smS_Label10.Name = "smS_Label10";
            this.smS_Label10.Size = new System.Drawing.Size(40, 12);
            this.smS_Label10.TabIndex = 53;
            this.smS_Label10.Text = "SKUCD";
            // 
            // lbl1
            // 
            this.lbl1.AutoSize = true;
            this.lbl1.BackColor = System.Drawing.Color.Transparent;
            this.lbl1.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.lbl1.Location = new System.Drawing.Point(605, 37);
            this.lbl1.Name = "lbl1";
            this.lbl1.Size = new System.Drawing.Size(57, 12);
            this.lbl1.TabIndex = 52;
            this.lbl1.Text = "カタログ";
            // 
            // smS_Label8
            // 
            this.smS_Label8.AutoSize = true;
            this.smS_Label8.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label8.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label8.Location = new System.Drawing.Point(618, 10);
            this.smS_Label8.Name = "smS_Label8";
            this.smS_Label8.Size = new System.Drawing.Size(44, 12);
            this.smS_Label8.TabIndex = 51;
            this.smS_Label8.Text = "取込日";
            // 
            // rdoMakerItemCD
            // 
            this.rdoMakerItemCD.AutoSize = true;
            this.rdoMakerItemCD.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rdoMakerItemCD.Location = new System.Drawing.Point(262, 137);
            this.rdoMakerItemCD.Name = "rdoMakerItemCD";
            this.rdoMakerItemCD.Size = new System.Drawing.Size(115, 16);
            this.rdoMakerItemCD.TabIndex = 12;
            this.rdoMakerItemCD.Text = "メーカー商品CD";
            this.rdoMakerItemCD.UseVisualStyleBackColor = true;
            // 
            // rdoItem
            // 
            this.rdoItem.AutoSize = true;
            this.rdoItem.Checked = true;
            this.rdoItem.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rdoItem.Location = new System.Drawing.Point(205, 137);
            this.rdoItem.Name = "rdoItem";
            this.rdoItem.Size = new System.Drawing.Size(51, 16);
            this.rdoItem.TabIndex = 11;
            this.rdoItem.TabStop = true;
            this.rdoItem.Text = "ITEM";
            this.rdoItem.UseVisualStyleBackColor = true;
            // 
            // chkSearch
            // 
            this.chkSearch.AutoSize = true;
            this.chkSearch.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkSearch.Location = new System.Drawing.Point(136, 137);
            this.chkSearch.Name = "chkSearch";
            this.chkSearch.Size = new System.Drawing.Size(57, 16);
            this.chkSearch.TabIndex = 11;
            this.chkSearch.TabStop = false;
            this.chkSearch.Text = " する";
            this.chkSearch.UseVisualStyleBackColor = true;
            // 
            // ucJANCD1
            // 
            this.ucJANCD1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            this.ucJANCD1.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ucJANCD1.Location = new System.Drawing.Point(132, 107);
            this.ucJANCD1.Margin = new System.Windows.Forms.Padding(0);
            this.ucJANCD1.Name = "ucJANCD1";
            this.ucJANCD1.Size = new System.Drawing.Size(980, 23);
            this.ucJANCD1.TabIndex = 9;
            this.ucJANCD1.UC_Code = "";
            this.ucJANCD1.UC_Code_Width = 940;
            this.ucJANCD1.UC_Exists = false;
            this.ucJANCD1.UC_Flag = 5;
            this.ucJANCD1.UC_IsRequired = false;
            this.ucJANCD1.UC_Name = "";
            this.ucJANCD1.UC_Name_Width = 281;
            this.ucJANCD1.UC_NameVisible = false;
            this.ucJANCD1.UC_SearchEnable = true;
            this.ucJANCD1.UC_Type = SMS.CustomControls.UC_Search.Type.MultiJANCD;
            this.ucJANCD1.Enter += new System.EventHandler(this.uc_Enter);
            this.ucJANCD1.Leave += new System.EventHandler(this.uc_Leave);
            // 
            // ucMakerItem
            // 
            this.ucMakerItem.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            this.ucMakerItem.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ucMakerItem.Location = new System.Drawing.Point(132, 82);
            this.ucMakerItem.Margin = new System.Windows.Forms.Padding(0);
            this.ucMakerItem.Name = "ucMakerItem";
            this.ucMakerItem.Size = new System.Drawing.Size(255, 23);
            this.ucMakerItem.TabIndex = 8;
            this.ucMakerItem.UC_Code = "";
            this.ucMakerItem.UC_Code_Width = 220;
            this.ucMakerItem.UC_Exists = false;
            this.ucMakerItem.UC_Flag = 2;
            this.ucMakerItem.UC_IsRequired = false;
            this.ucMakerItem.UC_Name = "";
            this.ucMakerItem.UC_Name_Width = 281;
            this.ucMakerItem.UC_NameVisible = false;
            this.ucMakerItem.UC_SearchEnable = true;
            this.ucMakerItem.UC_Type = SMS.CustomControls.UC_Search.Type.MakerShohinCD;
            this.ucMakerItem.Enter += new System.EventHandler(this.uc_Enter);
            this.ucMakerItem.Leave += new System.EventHandler(this.uc_Leave);
            // 
            // ucITEM
            // 
            this.ucITEM.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            this.ucITEM.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ucITEM.Location = new System.Drawing.Point(132, 57);
            this.ucITEM.Margin = new System.Windows.Forms.Padding(0);
            this.ucITEM.Name = "ucITEM";
            this.ucITEM.Size = new System.Drawing.Size(255, 23);
            this.ucITEM.TabIndex = 6;
            this.ucITEM.UC_Code = "";
            this.ucITEM.UC_Code_Width = 220;
            this.ucITEM.UC_Exists = false;
            this.ucITEM.UC_Flag = 1;
            this.ucITEM.UC_IsRequired = false;
            this.ucITEM.UC_Name = "";
            this.ucITEM.UC_Name_Width = 281;
            this.ucITEM.UC_NameVisible = false;
            this.ucITEM.UC_SearchEnable = true;
            this.ucITEM.UC_Type = SMS.CustomControls.UC_Search.Type.ITEM;
            this.ucITEM.Enter += new System.EventHandler(this.uc_Enter);
            this.ucITEM.Leave += new System.EventHandler(this.uc_Leave);
            // 
            // smS_Label7
            // 
            this.smS_Label7.AutoSize = true;
            this.smS_Label7.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label7.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label7.Location = new System.Drawing.Point(71, 139);
            this.smS_Label7.Name = "smS_Label7";
            this.smS_Label7.Size = new System.Drawing.Size(57, 12);
            this.smS_Label7.TabIndex = 42;
            this.smS_Label7.Text = "関連検索";
            // 
            // smS_Label6
            // 
            this.smS_Label6.AutoSize = true;
            this.smS_Label6.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label6.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label6.Location = new System.Drawing.Point(88, 113);
            this.smS_Label6.Name = "smS_Label6";
            this.smS_Label6.Size = new System.Drawing.Size(40, 12);
            this.smS_Label6.TabIndex = 41;
            this.smS_Label6.Text = "JANCD";
            // 
            // smS_Label5
            // 
            this.smS_Label5.AutoSize = true;
            this.smS_Label5.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label5.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label5.Location = new System.Drawing.Point(32, 88);
            this.smS_Label5.Name = "smS_Label5";
            this.smS_Label5.Size = new System.Drawing.Size(97, 12);
            this.smS_Label5.TabIndex = 40;
            this.smS_Label5.Text = "メーカー商品CD";
            // 
            // smS_Label4
            // 
            this.smS_Label4.AutoSize = true;
            this.smS_Label4.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label4.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label4.Location = new System.Drawing.Point(95, 63);
            this.smS_Label4.Name = "smS_Label4";
            this.smS_Label4.Size = new System.Drawing.Size(33, 12);
            this.smS_Label4.TabIndex = 39;
            this.smS_Label4.Text = "ITEM";
            // 
            // smS_Label3
            // 
            this.smS_Label3.AutoSize = true;
            this.smS_Label3.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label3.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label3.Location = new System.Drawing.Point(72, 36);
            this.smS_Label3.Name = "smS_Label3";
            this.smS_Label3.Size = new System.Drawing.Size(57, 12);
            this.smS_Label3.TabIndex = 38;
            this.smS_Label3.Text = "ブランド";
            // 
            // smS_Label2
            // 
            this.smS_Label2.AutoSize = true;
            this.smS_Label2.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label2.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label2.Location = new System.Drawing.Point(19, 11);
            this.smS_Label2.Name = "smS_Label2";
            this.smS_Label2.Size = new System.Drawing.Size(110, 12);
            this.smS_Label2.TabIndex = 37;
            this.smS_Label2.Text = "仕入先(データ元)";
            // 
            // smS_Label16
            // 
            this.smS_Label16.AutoSize = true;
            this.smS_Label16.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label16.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label16.Location = new System.Drawing.Point(38, 255);
            this.smS_Label16.Name = "smS_Label16";
            this.smS_Label16.Size = new System.Drawing.Size(83, 12);
            this.smS_Label16.TabIndex = 73;
            this.smS_Label16.Text = "【取込明細】";
            // 
            // dgvPSK0104I
            // 
            this.dgvPSK0104I.AllowUserToAddRows = false;
            this.dgvPSK0104I.AllowUserToDeleteRows = false;
            this.dgvPSK0104I.AllowUserToResizeRows = false;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(221)))), ((int)(((byte)(235)))), ((int)(((byte)(247)))));
            this.dgvPSK0104I.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dgvPSK0104I.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            dataGridViewCellStyle2.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgvPSK0104I.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle2;
            this.dgvPSK0104I.ColumnHeadersHeight = 25;
            this.dgvPSK0104I.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colNo,
            this.colFixedDateTime,
            this.colJANCD,
            this.colItemName,
            this.colToriKomiDateTime,
            this.colToriKomiNo,
            this.colKakuteiZaikoSu,
            this.colNouki,
            this.colTokukiJikou,
            this.colAdminCD,
            this.coluchiwake});
            dataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle7.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle7.Font = new System.Drawing.Font("MS Gothic", 9F);
            dataGridViewCellStyle7.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle7.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle7.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgvPSK0104I.DefaultCellStyle = dataGridViewCellStyle7;
            this.dgvPSK0104I.EnableHeadersVisualStyles = false;
            this.dgvPSK0104I.GridColor = System.Drawing.Color.DimGray;
            this.dgvPSK0104I.IsCellValid = false;
            this.dgvPSK0104I.Location = new System.Drawing.Point(44, 270);
            this.dgvPSK0104I.Name = "dgvPSK0104I";
            this.dgvPSK0104I.RowHeadersVisible = false;
            this.dgvPSK0104I.Size = new System.Drawing.Size(1360, 580);
            this.dgvPSK0104I.TabIndex = 19;
            this.dgvPSK0104I.CellEnter += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvPSK0104I_CellEnter);
            this.dgvPSK0104I.CellFormatting += new System.Windows.Forms.DataGridViewCellFormattingEventHandler(this.dgvPSK0104I_CellFormatting);
            // 
            // colNo
            // 
            this.colNo.DataPropertyName = "No";
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            this.colNo.DefaultCellStyle = dataGridViewCellStyle3;
            this.colNo.HeaderText = "No";
            this.colNo.Name = "colNo";
            this.colNo.ReadOnly = true;
            this.colNo.Width = 50;
            // 
            // colFixedDateTime
            // 
            this.colFixedDateTime.DataPropertyName = "FixedDateTime";
            this.colFixedDateTime.HeaderText = "FixedDateTime";
            this.colFixedDateTime.Name = "colFixedDateTime";
            this.colFixedDateTime.ReadOnly = true;
            this.colFixedDateTime.Visible = false;
            // 
            // colJANCD
            // 
            this.colJANCD.DataPropertyName = "JANCD";
            this.colJANCD.HeaderText = "JANCD";
            this.colJANCD.Name = "colJANCD";
            this.colJANCD.ReadOnly = true;
            this.colJANCD.Width = 90;
            // 
            // colItemName
            // 
            this.colItemName.DataPropertyName = "vm_hanbai_shohin";
            this.colItemName.HeaderText = "商品名";
            this.colItemName.Name = "colItemName";
            this.colItemName.ReadOnly = true;
            this.colItemName.Width = 500;
            // 
            // colToriKomiDateTime
            // 
            this.colToriKomiDateTime.CheckFormat = "yyyy/MM/dd hh:mm";
            this.colToriKomiDateTime.DataPropertyName = "InportDateTime";
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this.colToriKomiDateTime.DefaultCellStyle = dataGridViewCellStyle4;
            this.colToriKomiDateTime.HeaderText = "取込日時";
            this.colToriKomiDateTime.Mask = "0000/00/00 00:00";
            this.colToriKomiDateTime.Name = "colToriKomiDateTime";
            this.colToriKomiDateTime.ReadOnly = true;
            this.colToriKomiDateTime.Width = 150;
            // 
            // colToriKomiNo
            // 
            this.colToriKomiNo.DataPropertyName = "InportSu";
            this.colToriKomiNo.DecimalPlace = ((byte)(0));
            dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            this.colToriKomiNo.DefaultCellStyle = dataGridViewCellStyle5;
            this.colToriKomiNo.HeaderText = "取込在庫数";
            this.colToriKomiNo.MaxInputLength = 32767;
            this.colToriKomiNo.Name = "colToriKomiNo";
            this.colToriKomiNo.ReadOnly = true;
            this.colToriKomiNo.UseMinus = false;
            this.colToriKomiNo.UseThousandSeparator = true;
            this.colToriKomiNo.Width = 90;
            // 
            // colKakuteiZaikoSu
            // 
            this.colKakuteiZaikoSu.DataPropertyName = "UpdateSu";
            this.colKakuteiZaikoSu.DecimalPlace = ((byte)(0));
            dataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            this.colKakuteiZaikoSu.DefaultCellStyle = dataGridViewCellStyle6;
            this.colKakuteiZaikoSu.HeaderText = "確定在庫数";
            this.colKakuteiZaikoSu.MaxInputLength = 5;
            this.colKakuteiZaikoSu.Name = "colKakuteiZaikoSu";
            this.colKakuteiZaikoSu.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.colKakuteiZaikoSu.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.colKakuteiZaikoSu.UseMinus = false;
            this.colKakuteiZaikoSu.UseThousandSeparator = true;
            this.colKakuteiZaikoSu.Width = 90;
            // 
            // colNouki
            // 
            this.colNouki.DataPropertyName = "HenkanNouki";
            this.colNouki.HeaderText = "納期";
            this.colNouki.MaxInputLength = 30;
            this.colNouki.Name = "colNouki";
            this.colNouki.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.colNouki.Width = 200;
            // 
            // colTokukiJikou
            // 
            this.colTokukiJikou.DataPropertyName = "dgvMsg";
            this.colTokukiJikou.HeaderText = "特記事項";
            this.colTokukiJikou.Name = "colTokukiJikou";
            this.colTokukiJikou.ReadOnly = true;
            this.colTokukiJikou.Width = 170;
            // 
            // colAdminCD
            // 
            this.colAdminCD.DataPropertyName = "AdminCD";
            this.colAdminCD.HeaderText = "AdminCD";
            this.colAdminCD.Name = "colAdminCD";
            this.colAdminCD.Visible = false;
            // 
            // coluchiwake
            // 
            this.coluchiwake.DataPropertyName = "nf_uchiwake";
            this.coluchiwake.HeaderText = "uchiwake";
            this.coluchiwake.Name = "coluchiwake";
            this.coluchiwake.Visible = false;
            // 
            // smS_Label17
            // 
            this.smS_Label17.AutoSize = true;
            this.smS_Label17.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label17.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label17.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(55)))), ((int)(((byte)(86)))), ((int)(((byte)(35)))));
            this.smS_Label17.Location = new System.Drawing.Point(392, 140);
            this.smS_Label17.Name = "smS_Label17";
            this.smS_Label17.Size = new System.Drawing.Size(425, 12);
            this.smS_Label17.TabIndex = 14;
            this.smS_Label17.Text = "検索された商品と同じITEM(メーカー商品CD)の商品も同時に表示する。";
            // 
            // smS_Label18
            // 
            this.smS_Label18.BackColor = System.Drawing.Color.Black;
            this.smS_Label18.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.smS_Label18.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label18.Location = new System.Drawing.Point(10, 178);
            this.smS_Label18.Name = "smS_Label18";
            this.smS_Label18.Size = new System.Drawing.Size(1350, 2);
            this.smS_Label18.TabIndex = 76;
            // 
            // lblKakuteiMode
            // 
            this.lblKakuteiMode.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(176)))), ((int)(((byte)(240)))));
            this.lblKakuteiMode.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblKakuteiMode.Location = new System.Drawing.Point(167, 218);
            this.lblKakuteiMode.Name = "lblKakuteiMode";
            this.lblKakuteiMode.Size = new System.Drawing.Size(94, 32);
            this.lblKakuteiMode.TabIndex = 2;
            this.lblKakuteiMode.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // smS_Label19
            // 
            this.smS_Label19.AutoSize = true;
            this.smS_Label19.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label19.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label19.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.smS_Label19.Location = new System.Drawing.Point(144, 199);
            this.smS_Label19.Name = "smS_Label19";
            this.smS_Label19.Size = new System.Drawing.Size(430, 12);
            this.smS_Label19.TabIndex = 79;
            this.smS_Label19.Text = "この仕入先(データ元)からの在庫情報を確定し、SKSへの送信対象とする";
            // 
            // txtZaikoSu
            // 
            this.txtZaikoSu.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtZaikoSu.Decimalplace = ((byte)(0));
            this.txtZaikoSu.Location = new System.Drawing.Point(667, 222);
            this.txtZaikoSu.MaxLength = 4;
            this.txtZaikoSu.Name = "txtZaikoSu";
            this.txtZaikoSu.Size = new System.Drawing.Size(60, 19);
            this.txtZaikoSu.TabIndex = 14;
            this.txtZaikoSu.txt_Type = SMS.CustomControls.SMS_TextBox.Type.Decimal;
            this.txtZaikoSu.UseMinus = false;
            this.txtZaikoSu.UseThousandSeparator = true;
            // 
            // txtInportDate
            // 
            this.txtInportDate.CKM_Reqired = false;
            this.txtInportDate.CKM_Type = CKM_Controls.CKM_TextBox.Type.Date;
            this.txtInportDate.DecimalPlace = 0;
            this.txtInportDate.Location = new System.Drawing.Point(668, 6);
            this.txtInportDate.Name = "txtInportDate";
            this.txtInportDate.Size = new System.Drawing.Size(100, 19);
            this.txtInportDate.TabIndex = 1;
            this.txtInportDate.UseMinus = true;
            this.txtInportDate.UseThousandSeparator = true;
            // 
            // cboNendo
            // 
            this.cboNendo.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Append;
            this.cboNendo.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cboNendo.Cbo_Type = SMS.CustomControls.SMS_ComboBox.cbo_Type.m_nendo;
            this.cboNendo.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cboNendo.FormattingEnabled = true;
            this.cboNendo.Location = new System.Drawing.Point(965, 33);
            this.cboNendo.Name = "cboNendo";
            this.cboNendo.Size = new System.Drawing.Size(121, 20);
            this.cboNendo.TabIndex = 4;
            // 
            // cboSeason
            // 
            this.cboSeason.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Append;
            this.cboSeason.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cboSeason.Cbo_Type = SMS.CustomControls.SMS_ComboBox.cbo_Type.m_season;
            this.cboSeason.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cboSeason.FormattingEnabled = true;
            this.cboSeason.Location = new System.Drawing.Point(1166, 33);
            this.cboSeason.Name = "cboSeason";
            this.cboSeason.Size = new System.Drawing.Size(121, 20);
            this.cboSeason.TabIndex = 5;
            // 
            // ucSupplier
            // 
            this.ucSupplier.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            this.ucSupplier.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ucSupplier.Location = new System.Drawing.Point(132, 5);
            this.ucSupplier.Margin = new System.Windows.Forms.Padding(0);
            this.ucSupplier.Name = "ucSupplier";
            this.ucSupplier.Size = new System.Drawing.Size(450, 25);
            this.ucSupplier.TabIndex = 0;
            this.ucSupplier.UC_Code = "";
            this.ucSupplier.UC_Code_Width = 65;
            this.ucSupplier.UC_Exists = false;
            this.ucSupplier.UC_Flag = 0;
            this.ucSupplier.UC_IsRequired = true;
            this.ucSupplier.UC_Name = "";
            this.ucSupplier.UC_Name_Width = 330;
            this.ucSupplier.UC_NameVisible = true;
            this.ucSupplier.UC_SearchEnable = true;
            this.ucSupplier.UC_Type = SMS.CustomControls.UC_Search.Type.Maker;
            this.ucSupplier.Enter += new System.EventHandler(this.uc_Enter);
            this.ucSupplier.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ucSupplier_KeyDown);
            this.ucSupplier.Leave += new System.EventHandler(this.uc_Leave);
            // 
            // ucBrand
            // 
            this.ucBrand.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            this.ucBrand.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ucBrand.Location = new System.Drawing.Point(132, 30);
            this.ucBrand.Margin = new System.Windows.Forms.Padding(0);
            this.ucBrand.Name = "ucBrand";
            this.ucBrand.Size = new System.Drawing.Size(500, 23);
            this.ucBrand.TabIndex = 2;
            this.ucBrand.UC_Code = "";
            this.ucBrand.UC_Code_Width = 65;
            this.ucBrand.UC_Exists = false;
            this.ucBrand.UC_Flag = 1;
            this.ucBrand.UC_IsRequired = false;
            this.ucBrand.UC_Name = "";
            this.ucBrand.UC_Name_Width = 330;
            this.ucBrand.UC_NameVisible = true;
            this.ucBrand.UC_SearchEnable = true;
            this.ucBrand.UC_Type = SMS.CustomControls.UC_Search.Type.Brand;
            this.ucBrand.Enter += new System.EventHandler(this.uc_Enter);
            this.ucBrand.Leave += new System.EventHandler(this.uc_Leave);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.smS_Label2);
            this.panel1.Controls.Add(this.ucBrand);
            this.panel1.Controls.Add(this.smS_Label3);
            this.panel1.Controls.Add(this.ucSupplier);
            this.panel1.Controls.Add(this.smS_Label4);
            this.panel1.Controls.Add(this.cboSeason);
            this.panel1.Controls.Add(this.smS_Label5);
            this.panel1.Controls.Add(this.cboNendo);
            this.panel1.Controls.Add(this.smS_Label6);
            this.panel1.Controls.Add(this.txtInportDate);
            this.panel1.Controls.Add(this.smS_Label7);
            this.panel1.Controls.Add(this.ucITEM);
            this.panel1.Controls.Add(this.ucMakerItem);
            this.panel1.Controls.Add(this.ucJANCD1);
            this.panel1.Controls.Add(this.chkSearch);
            this.panel1.Controls.Add(this.smS_Label17);
            this.panel1.Controls.Add(this.rdoItem);
            this.panel1.Controls.Add(this.rdoMakerItemCD);
            this.panel1.Controls.Add(this.smS_Label8);
            this.panel1.Controls.Add(this.lbl1);
            this.panel1.Controls.Add(this.smS_Label10);
            this.panel1.Controls.Add(this.smS_Label11);
            this.panel1.Controls.Add(this.chkJANCDcut);
            this.panel1.Controls.Add(this.chkZaikoSuZero);
            this.panel1.Controls.Add(this.ucSKUCD);
            this.panel1.Controls.Add(this.btnDisplay);
            this.panel1.Controls.Add(this.smS_Label12);
            this.panel1.Controls.Add(this.txtCatalogue);
            this.panel1.Controls.Add(this.smS_Label13);
            this.panel1.Location = new System.Drawing.Point(6, 7);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1400, 165);
            this.panel1.TabIndex = 80;
            // 
            // frmPSKS0104I
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1884, 951);
            this.F5_Text = "ｷｬﾝｾﾙ(F5)";
            this.KeyPreview = true;
            this.Location = new System.Drawing.Point(0, 0);
            this.Name = "frmPSKS0104I";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "PSKS0104I";
            this.Load += new System.EventHandler(this.frmPSKS0104I_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmPSKS0104I_KeyDown);
            this.panelBody.ResumeLayout(false);
            this.panelBody.PerformLayout();
            this.panelTitleLeft.ResumeLayout(false);
            this.panelTitle.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvPSK0104I)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private CustomControls.SMS_Button btnAdd;
        private CustomControls.SMS_Label smS_Label15;
        private CustomControls.UC_Search ucJANCD2;
        private CustomControls.SMS_Label smS_Label14;
        private CustomControls.SMS_Button btnMiKakutei;
        private CustomControls.SMS_Button btnKakutei;
        private CustomControls.SMS_Button btnDisplay;
        private CustomControls.SMS_TextBox txtCatalogue;
        private CustomControls.SMS_Label smS_Label13;
        private CustomControls.SMS_Label smS_Label12;
        private CustomControls.UC_Search ucSKUCD;
        private System.Windows.Forms.CheckBox chkZaikoSuZero;
        private System.Windows.Forms.CheckBox chkJANCDcut;
        private CustomControls.SMS_Label smS_Label11;
        private CustomControls.SMS_Label smS_Label10;
        private CustomControls.SMS_Label lbl1;
        private CustomControls.SMS_Label smS_Label8;
        private System.Windows.Forms.RadioButton rdoMakerItemCD;
        private System.Windows.Forms.RadioButton rdoItem;
        private System.Windows.Forms.CheckBox chkSearch;
        private CustomControls.UC_Search ucJANCD1;
        private CustomControls.UC_Search ucMakerItem;
        private CustomControls.UC_Search ucITEM;
        private CustomControls.SMS_Label smS_Label7;
        private CustomControls.SMS_Label smS_Label6;
        private CustomControls.SMS_Label smS_Label5;
        private CustomControls.SMS_Label smS_Label4;
        private CustomControls.SMS_Label smS_Label3;
        private CustomControls.SMS_Label smS_Label2;
        private CustomControls.SMS_Label smS_Label16;
        private SMS_GridView dgvPSK0104I;
        private CustomControls.SMS_Label smS_Label17;
        private CustomControls.SMS_Label smS_Label18;
        private CustomControls.SMS_TextBox txtZaikoSu;
        private CustomControls.SMS_Label smS_Label19;
        public System.Windows.Forms.Label lblKakuteiMode;
        private CKM_Controls.CKM_TextBox txtInportDate;
        private System.Windows.Forms.DataGridViewTextBoxColumn colNo;
        private System.Windows.Forms.DataGridViewTextBoxColumn colFixedDateTime;
        private System.Windows.Forms.DataGridViewTextBoxColumn colJANCD;
        private System.Windows.Forms.DataGridViewTextBoxColumn colItemName;
        private CustomControls.DataGridViewMaskedTextColumn colToriKomiDateTime;
        private CustomControls.DataGridViewDecimalColumn colToriKomiNo;
        private CustomControls.DataGridViewDecimalColumn colKakuteiZaikoSu;
        private System.Windows.Forms.DataGridViewTextBoxColumn colNouki;
        private System.Windows.Forms.DataGridViewTextBoxColumn colTokukiJikou;
        private System.Windows.Forms.DataGridViewTextBoxColumn colAdminCD;
        private System.Windows.Forms.DataGridViewTextBoxColumn coluchiwake;
        private CustomControls.SMS_ComboBox cboSeason;
        private CustomControls.SMS_ComboBox cboNendo;
        private CustomControls.UC_Search ucSupplier;
        private CustomControls.UC_Search ucBrand;
        private System.Windows.Forms.Panel panel1;
    }
}