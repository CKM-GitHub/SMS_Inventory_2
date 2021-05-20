namespace SMS.Inventory
{
    partial class frmPSKS0101M
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle23 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle24 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle26 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle25 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle19 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle20 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle22 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle21 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle14 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle15 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle18 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle16 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle17 = new System.Windows.Forms.DataGridViewCellStyle();
            this.ucPattern = new SMS.CustomControls.UC_Search();
            this.txtPatternName = new SMS.CustomControls.SMS_TextBox();
            this.smS_Label8 = new SMS.CustomControls.SMS_Label();
            this.smS_Label9 = new SMS.CustomControls.SMS_Label();
            this.smS_Label11 = new SMS.CustomControls.SMS_Label();
            this.smS_Label10 = new SMS.CustomControls.SMS_Label();
            this.dgvdelivery = new SMS.SMS_GridView();
            this.colNum = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colNouki = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colHenkanNouki = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dgvKamoku = new SMS.SMS_GridView();
            this.colNo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colZokuSei = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.smS_Label7 = new SMS.CustomControls.SMS_Label();
            this.ucCopyPattern = new SMS.CustomControls.UC_Search();
            this.smS_GroupBox1 = new SMS.CustomControls.SMS_GroupBox();
            this.dgvSuuryoHenKan = new SMS.SMS_GridView();
            this.colSEQ = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colMakerSuuryo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colConversionSuuryo = new SMS.CustomControls.DataGridViewDecimalColumn();
            this.smS_Label2 = new SMS.CustomControls.SMS_Label();
            this.panelBody.SuspendLayout();
            this.panelTitleLeft.SuspendLayout();
            this.panelTitle.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvdelivery)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvKamoku)).BeginInit();
            this.smS_GroupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvSuuryoHenKan)).BeginInit();
            this.SuspendLayout();
            // 
            // panelBody
            // 
            this.panelBody.Controls.Add(this.dgvSuuryoHenKan);
            this.panelBody.Controls.Add(this.smS_Label2);
            this.panelBody.Controls.Add(this.smS_GroupBox1);
            this.panelBody.Controls.Add(this.smS_Label9);
            this.panelBody.Controls.Add(this.smS_Label8);
            this.panelBody.Controls.Add(this.txtPatternName);
            this.panelBody.Controls.Add(this.dgvKamoku);
            this.panelBody.Controls.Add(this.ucPattern);
            this.panelBody.Controls.Add(this.dgvdelivery);
            this.panelBody.Controls.Add(this.smS_Label10);
            this.panelBody.Controls.Add(this.smS_Label11);
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
            // ucPattern
            // 
            this.ucPattern.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            this.ucPattern.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ucPattern.Location = new System.Drawing.Point(106, 7);
            this.ucPattern.Margin = new System.Windows.Forms.Padding(0);
            this.ucPattern.Name = "ucPattern";
            this.ucPattern.Size = new System.Drawing.Size(83, 23);
            this.ucPattern.TabIndex = 1;
            this.ucPattern.UC_Code = "";
            this.ucPattern.UC_Code_Width = 30;
            this.ucPattern.UC_Exists = false;
            this.ucPattern.UC_Flag = 0;
            this.ucPattern.UC_IsRequired = false;
            this.ucPattern.UC_Name = "";
            this.ucPattern.UC_Name_Width = 270;
            this.ucPattern.UC_NameVisible = false;
            this.ucPattern.UC_SearchEnable = true;
            this.ucPattern.UC_Type = SMS.CustomControls.UC_Search.Type.Pattern;
            this.ucPattern.Enter += new System.EventHandler(this.ucPattern_Enter);
            this.ucPattern.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ucPattern_KeyDown);
            this.ucPattern.Leave += new System.EventHandler(this.uc_Leave);
            // 
            // txtPatternName
            // 
            this.txtPatternName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtPatternName.Decimalplace = ((byte)(0));
            this.txtPatternName.Location = new System.Drawing.Point(106, 31);
            this.txtPatternName.MaxLength = 20;
            this.txtPatternName.Name = "txtPatternName";
            this.txtPatternName.Size = new System.Drawing.Size(270, 19);
            this.txtPatternName.TabIndex = 2;
            this.txtPatternName.txt_Type = SMS.CustomControls.SMS_TextBox.Type.Hiragana;
            this.txtPatternName.UseMinus = false;
            this.txtPatternName.UseThousandSeparator = true;
            this.txtPatternName.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtPatternName_KeyDown);
            // 
            // smS_Label8
            // 
            this.smS_Label8.AutoSize = true;
            this.smS_Label8.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label8.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label8.Location = new System.Drawing.Point(25, 35);
            this.smS_Label8.Name = "smS_Label8";
            this.smS_Label8.Size = new System.Drawing.Size(70, 12);
            this.smS_Label8.TabIndex = 46;
            this.smS_Label8.Text = "パターン名";
            // 
            // smS_Label9
            // 
            this.smS_Label9.AutoSize = true;
            this.smS_Label9.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label9.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label9.Location = new System.Drawing.Point(12, 13);
            this.smS_Label9.Name = "smS_Label9";
            this.smS_Label9.Size = new System.Drawing.Size(83, 12);
            this.smS_Label9.TabIndex = 50;
            this.smS_Label9.Text = "取込パターン";
            // 
            // smS_Label11
            // 
            this.smS_Label11.AutoSize = true;
            this.smS_Label11.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label11.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label11.Location = new System.Drawing.Point(560, 72);
            this.smS_Label11.Name = "smS_Label11";
            this.smS_Label11.Size = new System.Drawing.Size(57, 12);
            this.smS_Label11.TabIndex = 64;
            this.smS_Label11.Text = "納期変換";
            // 
            // smS_Label10
            // 
            this.smS_Label10.AutoSize = true;
            this.smS_Label10.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label10.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label10.Location = new System.Drawing.Point(41, 72);
            this.smS_Label10.Name = "smS_Label10";
            this.smS_Label10.Size = new System.Drawing.Size(57, 12);
            this.smS_Label10.TabIndex = 63;
            this.smS_Label10.Text = "項目定義";
            // 
            // dgvdelivery
            // 
            this.dgvdelivery.AllowUserToAddRows = false;
            this.dgvdelivery.AllowUserToDeleteRows = false;
            this.dgvdelivery.AllowUserToResizeRows = false;
            dataGridViewCellStyle23.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(221)))), ((int)(((byte)(235)))), ((int)(((byte)(247)))));
            this.dgvdelivery.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle23;
            this.dgvdelivery.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            dataGridViewCellStyle24.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle24.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            dataGridViewCellStyle24.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle24.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle24.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle24.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle24.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgvdelivery.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle24;
            this.dgvdelivery.ColumnHeadersHeight = 25;
            this.dgvdelivery.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.dgvdelivery.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colNum,
            this.colNouki,
            this.colHenkanNouki});
            dataGridViewCellStyle26.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle26.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle26.Font = new System.Drawing.Font("MS Gothic", 9F);
            dataGridViewCellStyle26.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle26.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle26.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle26.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgvdelivery.DefaultCellStyle = dataGridViewCellStyle26;
            this.dgvdelivery.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter;
            this.dgvdelivery.EnableHeadersVisualStyles = false;
            this.dgvdelivery.GridColor = System.Drawing.Color.DimGray;
            this.dgvdelivery.IsCellValid = false;
            this.dgvdelivery.Location = new System.Drawing.Point(576, 87);
            this.dgvdelivery.Name = "dgvdelivery";
            this.dgvdelivery.RowHeadersVisible = false;
            this.dgvdelivery.Size = new System.Drawing.Size(470, 690);
            this.dgvdelivery.TabIndex = 5;
            this.dgvdelivery.CellEnter += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv_CellEnter);
            this.dgvdelivery.EditingControlShowing += new System.Windows.Forms.DataGridViewEditingControlShowingEventHandler(this.dgvdelivery_EditingControlShowing);
            // 
            // colNum
            // 
            this.colNum.DataPropertyName = "SEQ";
            dataGridViewCellStyle25.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            this.colNum.DefaultCellStyle = dataGridViewCellStyle25;
            this.colNum.HeaderText = "No";
            this.colNum.Name = "colNum";
            this.colNum.ReadOnly = true;
            this.colNum.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.colNum.Width = 50;
            // 
            // colNouki
            // 
            this.colNouki.DataPropertyName = "Maker";
            this.colNouki.HeaderText = "メーカー値";
            this.colNouki.MaxInputLength = 30;
            this.colNouki.Name = "colNouki";
            this.colNouki.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.colNouki.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.colNouki.Width = 200;
            // 
            // colHenkanNouki
            // 
            this.colHenkanNouki.DataPropertyName = "Conversion";
            this.colHenkanNouki.HeaderText = "変換値";
            this.colHenkanNouki.MaxInputLength = 30;
            this.colHenkanNouki.Name = "colHenkanNouki";
            this.colHenkanNouki.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.colHenkanNouki.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.colHenkanNouki.Width = 200;
            // 
            // dgvKamoku
            // 
            this.dgvKamoku.AllowUserToAddRows = false;
            this.dgvKamoku.AllowUserToDeleteRows = false;
            this.dgvKamoku.AllowUserToResizeRows = false;
            dataGridViewCellStyle19.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(221)))), ((int)(((byte)(235)))), ((int)(((byte)(247)))));
            this.dgvKamoku.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle19;
            this.dgvKamoku.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            dataGridViewCellStyle20.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle20.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            dataGridViewCellStyle20.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle20.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle20.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle20.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle20.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgvKamoku.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle20;
            this.dgvKamoku.ColumnHeadersHeight = 25;
            this.dgvKamoku.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.dgvKamoku.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colNo,
            this.colName,
            this.colZokuSei});
            dataGridViewCellStyle22.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle22.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle22.Font = new System.Drawing.Font("MS Gothic", 9F);
            dataGridViewCellStyle22.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle22.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle22.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle22.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgvKamoku.DefaultCellStyle = dataGridViewCellStyle22;
            this.dgvKamoku.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter;
            this.dgvKamoku.EnableHeadersVisualStyles = false;
            this.dgvKamoku.GridColor = System.Drawing.Color.DimGray;
            this.dgvKamoku.IsCellValid = false;
            this.dgvKamoku.Location = new System.Drawing.Point(56, 87);
            this.dgvKamoku.Name = "dgvKamoku";
            this.dgvKamoku.RowHeadersVisible = false;
            this.dgvKamoku.Size = new System.Drawing.Size(450, 690);
            this.dgvKamoku.TabIndex = 4;
            this.dgvKamoku.CellEnter += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv_CellEnter);
            this.dgvKamoku.EditingControlShowing += new System.Windows.Forms.DataGridViewEditingControlShowingEventHandler(this.dgvKamoku_EditingControlShowing);
            // 
            // colNo
            // 
            this.colNo.DataPropertyName = "SEQ";
            dataGridViewCellStyle21.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            this.colNo.DefaultCellStyle = dataGridViewCellStyle21;
            this.colNo.HeaderText = "No";
            this.colNo.Name = "colNo";
            this.colNo.ReadOnly = true;
            this.colNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.colNo.Width = 50;
            // 
            // colName
            // 
            this.colName.DataPropertyName = "ItemName";
            this.colName.HeaderText = "名前";
            this.colName.MaxInputLength = 30;
            this.colName.Name = "colName";
            this.colName.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.colName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.colName.Width = 200;
            // 
            // colZokuSei
            // 
            this.colZokuSei.DataPropertyName = "ItemProperty";
            this.colZokuSei.HeaderText = "属性";
            this.colZokuSei.Name = "colZokuSei";
            this.colZokuSei.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.colZokuSei.Width = 180;
            // 
            // smS_Label7
            // 
            this.smS_Label7.AutoSize = true;
            this.smS_Label7.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label7.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label7.Location = new System.Drawing.Point(44, 20);
            this.smS_Label7.Name = "smS_Label7";
            this.smS_Label7.Size = new System.Drawing.Size(83, 12);
            this.smS_Label7.TabIndex = 0;
            this.smS_Label7.Text = "取込パターン";
            // 
            // ucCopyPattern
            // 
            this.ucCopyPattern.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            this.ucCopyPattern.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ucCopyPattern.Location = new System.Drawing.Point(130, 14);
            this.ucCopyPattern.Margin = new System.Windows.Forms.Padding(0);
            this.ucCopyPattern.Name = "ucCopyPattern";
            this.ucCopyPattern.Size = new System.Drawing.Size(340, 23);
            this.ucCopyPattern.TabIndex = 0;
            this.ucCopyPattern.UC_Code = "";
            this.ucCopyPattern.UC_Code_Width = 30;
            this.ucCopyPattern.UC_Exists = false;
            this.ucCopyPattern.UC_Flag = 0;
            this.ucCopyPattern.UC_IsRequired = false;
            this.ucCopyPattern.UC_Name = "";
            this.ucCopyPattern.UC_Name_Width = 250;
            this.ucCopyPattern.UC_NameVisible = true;
            this.ucCopyPattern.UC_SearchEnable = true;
            this.ucCopyPattern.UC_Type = SMS.CustomControls.UC_Search.Type.Pattern;
            this.ucCopyPattern.Enter += new System.EventHandler(this.ucCopyPattern_Enter);
            this.ucCopyPattern.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ucCopyPattern_KeyDown);
            this.ucCopyPattern.Leave += new System.EventHandler(this.uc_Leave);
            // 
            // smS_GroupBox1
            // 
            this.smS_GroupBox1.Controls.Add(this.smS_Label7);
            this.smS_GroupBox1.Controls.Add(this.ucCopyPattern);
            this.smS_GroupBox1.Location = new System.Drawing.Point(577, 13);
            this.smS_GroupBox1.Name = "smS_GroupBox1";
            this.smS_GroupBox1.Size = new System.Drawing.Size(480, 50);
            this.smS_GroupBox1.TabIndex = 3;
            this.smS_GroupBox1.TabStop = false;
            this.smS_GroupBox1.Text = "複写元";
            // 
            // dgvSuuryoHenKan
            // 
            this.dgvSuuryoHenKan.AllowUserToAddRows = false;
            this.dgvSuuryoHenKan.AllowUserToDeleteRows = false;
            this.dgvSuuryoHenKan.AllowUserToResizeRows = false;
            dataGridViewCellStyle14.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(221)))), ((int)(((byte)(235)))), ((int)(((byte)(247)))));
            this.dgvSuuryoHenKan.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle14;
            this.dgvSuuryoHenKan.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            dataGridViewCellStyle15.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle15.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            dataGridViewCellStyle15.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle15.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle15.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle15.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle15.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgvSuuryoHenKan.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle15;
            this.dgvSuuryoHenKan.ColumnHeadersHeight = 25;
            this.dgvSuuryoHenKan.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.dgvSuuryoHenKan.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colSEQ,
            this.colMakerSuuryo,
            this.colConversionSuuryo});
            dataGridViewCellStyle18.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle18.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle18.Font = new System.Drawing.Font("MS Gothic", 9F);
            dataGridViewCellStyle18.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle18.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle18.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle18.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgvSuuryoHenKan.DefaultCellStyle = dataGridViewCellStyle18;
            this.dgvSuuryoHenKan.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter;
            this.dgvSuuryoHenKan.EnableHeadersVisualStyles = false;
            this.dgvSuuryoHenKan.GridColor = System.Drawing.Color.DimGray;
            this.dgvSuuryoHenKan.IsCellValid = false;
            this.dgvSuuryoHenKan.Location = new System.Drawing.Point(1116, 87);
            this.dgvSuuryoHenKan.Name = "dgvSuuryoHenKan";
            this.dgvSuuryoHenKan.RowHeadersVisible = false;
            this.dgvSuuryoHenKan.Size = new System.Drawing.Size(420, 690);
            this.dgvSuuryoHenKan.TabIndex = 65;
            this.dgvSuuryoHenKan.CellEnter += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv_CellEnter);
            // 
            // colSEQ
            // 
            this.colSEQ.DataPropertyName = "SEQSuuryo";
            dataGridViewCellStyle16.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            this.colSEQ.DefaultCellStyle = dataGridViewCellStyle16;
            this.colSEQ.HeaderText = "No";
            this.colSEQ.Name = "colSEQ";
            this.colSEQ.ReadOnly = true;
            this.colSEQ.Width = 50;
            // 
            // colMakerSuuryo
            // 
            this.colMakerSuuryo.DataPropertyName = "MakerSuuryo";
            this.colMakerSuuryo.HeaderText = "メーカー値";
            this.colMakerSuuryo.Name = "colMakerSuuryo";
            this.colMakerSuuryo.Width = 200;
            // 
            // colConversionSuuryo
            // 
            this.colConversionSuuryo.DataPropertyName = "ConversionSuuryo";
            this.colConversionSuuryo.DecimalPlace = ((byte)(0));
            dataGridViewCellStyle17.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            this.colConversionSuuryo.DefaultCellStyle = dataGridViewCellStyle17;
            this.colConversionSuuryo.HeaderText = "変換値";
            this.colConversionSuuryo.MaxInputLength = 32767;
            this.colConversionSuuryo.Name = "colConversionSuuryo";
            this.colConversionSuuryo.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.colConversionSuuryo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.colConversionSuuryo.UseMinus = false;
            this.colConversionSuuryo.UseThousandSeparator = true;
            this.colConversionSuuryo.Width = 150;
            // 
            // smS_Label2
            // 
            this.smS_Label2.AutoSize = true;
            this.smS_Label2.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label2.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label2.Location = new System.Drawing.Point(1100, 72);
            this.smS_Label2.Name = "smS_Label2";
            this.smS_Label2.Size = new System.Drawing.Size(57, 12);
            this.smS_Label2.TabIndex = 66;
            this.smS_Label2.Text = "数量変換";
            // 
            // frmPSKS0101M
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1884, 951);
            this.KeyPreview = true;
            this.Location = new System.Drawing.Point(0, 0);
            this.Name = "frmPSKS0101M";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "PSKS0101M";
            this.Load += new System.EventHandler(this.frmPSKS0101M_Load);
            this.panelBody.ResumeLayout(false);
            this.panelBody.PerformLayout();
            this.panelTitleLeft.ResumeLayout(false);
            this.panelTitle.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvdelivery)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvKamoku)).EndInit();
            this.smS_GroupBox1.ResumeLayout(false);
            this.smS_GroupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvSuuryoHenKan)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private CustomControls.SMS_Label smS_Label9;
        private CustomControls.UC_Search ucPattern;
        private CustomControls.SMS_TextBox txtPatternName;
        private CustomControls.SMS_Label smS_Label8;
        private CustomControls.SMS_Label smS_Label11;
        private CustomControls.SMS_Label smS_Label10;
        private SMS_GridView dgvdelivery;
        private SMS_GridView dgvKamoku;
        private CustomControls.SMS_GroupBox smS_GroupBox1;
        private CustomControls.SMS_Label smS_Label7;
        private CustomControls.UC_Search ucCopyPattern;
        private SMS_GridView dgvSuuryoHenKan;
        private CustomControls.SMS_Label smS_Label2;
        private System.Windows.Forms.DataGridViewTextBoxColumn colSEQ;
        private System.Windows.Forms.DataGridViewTextBoxColumn colMakerSuuryo;
        private CustomControls.DataGridViewDecimalColumn colConversionSuuryo;
        private System.Windows.Forms.DataGridViewTextBoxColumn colNo;
        private System.Windows.Forms.DataGridViewTextBoxColumn colName;
        private System.Windows.Forms.DataGridViewComboBoxColumn colZokuSei;
        private System.Windows.Forms.DataGridViewTextBoxColumn colNum;
        private System.Windows.Forms.DataGridViewTextBoxColumn colNouki;
        private System.Windows.Forms.DataGridViewTextBoxColumn colHenkanNouki;


    }
}