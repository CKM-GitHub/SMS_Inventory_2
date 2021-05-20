namespace SMS.Inventory
{
    partial class frmPSKS0102M
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle17 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle18 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle20 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle19 = new System.Windows.Forms.DataGridViewCellStyle();
            this.button11 = new SMS.CustomControls.SMS_Button();
            this.smS_Label2 = new SMS.CustomControls.SMS_Label();
            this.smS_Label3 = new SMS.CustomControls.SMS_Label();
            this.ucBrand = new SMS.CustomControls.UC_Search();
            this.ucMaker = new SMS.CustomControls.UC_Search();
            this.dgvM_MakerZaiko = new SMS.SMS_GridView();
            this.colNo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colMakerCd = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btnMaker = new System.Windows.Forms.DataGridViewButtonColumn();
            this.colMakerName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colBrandCD = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btnBrand = new System.Windows.Forms.DataGridViewButtonColumn();
            this.colBrandName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colData = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btnData = new System.Windows.Forms.DataGridViewButtonColumn();
            this.colDataName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colPatternCD = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btnPattern = new System.Windows.Forms.DataGridViewButtonColumn();
            this.colPatternName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colIrregularKBN = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.smS_Label7 = new SMS.CustomControls.SMS_Label();
            this.smS_Label4 = new SMS.CustomControls.SMS_Label();
            this.uc_DatasourceMaker = new SMS.CustomControls.UC_Search();
            this.panelBody.SuspendLayout();
            this.panelTitleLeft.SuspendLayout();
            this.panelTitle.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvM_MakerZaiko)).BeginInit();
            this.SuspendLayout();
            // 
            // panelBody
            // 
            this.panelBody.Controls.Add(this.smS_Label4);
            this.panelBody.Controls.Add(this.uc_DatasourceMaker);
            this.panelBody.Controls.Add(this.smS_Label7);
            this.panelBody.Controls.Add(this.dgvM_MakerZaiko);
            this.panelBody.Controls.Add(this.button11);
            this.panelBody.Controls.Add(this.smS_Label2);
            this.panelBody.Controls.Add(this.smS_Label3);
            this.panelBody.Controls.Add(this.ucBrand);
            this.panelBody.Controls.Add(this.ucMaker);
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
            // button11
            // 
            this.button11.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.button11.Cursor = System.Windows.Forms.Cursors.Hand;
            this.button11.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.button11.Location = new System.Drawing.Point(1561, 34);
            this.button11.Name = "button11";
            this.button11.Size = new System.Drawing.Size(100, 28);
            this.button11.TabIndex = 2;
            this.button11.Text = "表示(F11)";
            this.button11.UseVisualStyleBackColor = false;
            this.button11.Click += new System.EventHandler(this.button11_Click);
            // 
            // smS_Label2
            // 
            this.smS_Label2.AutoSize = true;
            this.smS_Label2.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label2.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label2.Location = new System.Drawing.Point(12, 13);
            this.smS_Label2.Name = "smS_Label2";
            this.smS_Label2.Size = new System.Drawing.Size(110, 12);
            this.smS_Label2.TabIndex = 3;
            this.smS_Label2.Text = "メーカー(仕入先)";
            // 
            // smS_Label3
            // 
            this.smS_Label3.AutoSize = true;
            this.smS_Label3.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label3.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label3.Location = new System.Drawing.Point(65, 43);
            this.smS_Label3.Name = "smS_Label3";
            this.smS_Label3.Size = new System.Drawing.Size(57, 12);
            this.smS_Label3.TabIndex = 19;
            this.smS_Label3.Text = "ブランド";
            // 
            // ucBrand
            // 
            this.ucBrand.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            this.ucBrand.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ucBrand.Location = new System.Drawing.Point(125, 37);
            this.ucBrand.Margin = new System.Windows.Forms.Padding(0);
            this.ucBrand.Name = "ucBrand";
            this.ucBrand.Size = new System.Drawing.Size(400, 25);
            this.ucBrand.TabIndex = 1;
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
            this.ucBrand.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ucBrand_KeyDown);
            this.ucBrand.Leave += new System.EventHandler(this.uc_Leave);
            // 
            // ucMaker
            // 
            this.ucMaker.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            this.ucMaker.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ucMaker.Location = new System.Drawing.Point(125, 7);
            this.ucMaker.Margin = new System.Windows.Forms.Padding(0);
            this.ucMaker.Name = "ucMaker";
            this.ucMaker.Size = new System.Drawing.Size(450, 25);
            this.ucMaker.TabIndex = 0;
            this.ucMaker.UC_Code = "";
            this.ucMaker.UC_Code_Width = 50;
            this.ucMaker.UC_Exists = false;
            this.ucMaker.UC_Flag = 0;
            this.ucMaker.UC_IsRequired = false;
            this.ucMaker.UC_Name = "";
            this.ucMaker.UC_Name_Width = 330;
            this.ucMaker.UC_NameVisible = true;
            this.ucMaker.UC_SearchEnable = true;
            this.ucMaker.UC_Type = SMS.CustomControls.UC_Search.Type.Maker;
            this.ucMaker.Enter += new System.EventHandler(this.uc_Enter);
            this.ucMaker.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ucMaker_KeyDown);
            this.ucMaker.Leave += new System.EventHandler(this.uc_Leave);
            // 
            // dgvM_MakerZaiko
            // 
            this.dgvM_MakerZaiko.AllowUserToAddRows = false;
            this.dgvM_MakerZaiko.AllowUserToDeleteRows = false;
            this.dgvM_MakerZaiko.AllowUserToResizeRows = false;
            dataGridViewCellStyle17.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(221)))), ((int)(((byte)(235)))), ((int)(((byte)(247)))));
            this.dgvM_MakerZaiko.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle17;
            this.dgvM_MakerZaiko.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            dataGridViewCellStyle18.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle18.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            dataGridViewCellStyle18.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle18.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle18.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle18.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle18.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgvM_MakerZaiko.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle18;
            this.dgvM_MakerZaiko.ColumnHeadersHeight = 25;
            this.dgvM_MakerZaiko.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colNo,
            this.Column1,
            this.Column2,
            this.colMakerCd,
            this.btnMaker,
            this.colMakerName,
            this.colBrandCD,
            this.btnBrand,
            this.colBrandName,
            this.colData,
            this.btnData,
            this.colDataName,
            this.colPatternCD,
            this.btnPattern,
            this.colPatternName,
            this.colIrregularKBN});
            dataGridViewCellStyle20.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle20.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle20.Font = new System.Drawing.Font("MS Gothic", 9F);
            dataGridViewCellStyle20.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle20.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle20.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle20.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgvM_MakerZaiko.DefaultCellStyle = dataGridViewCellStyle20;
            this.dgvM_MakerZaiko.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter;
            this.dgvM_MakerZaiko.EnableHeadersVisualStyles = false;
            this.dgvM_MakerZaiko.GridColor = System.Drawing.Color.DimGray;
            this.dgvM_MakerZaiko.IsCellValid = false;
            this.dgvM_MakerZaiko.Location = new System.Drawing.Point(12, 79);
            this.dgvM_MakerZaiko.Name = "dgvM_MakerZaiko";
            this.dgvM_MakerZaiko.RowHeadersVisible = false;
            this.dgvM_MakerZaiko.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.dgvM_MakerZaiko.Size = new System.Drawing.Size(1630, 685);
            this.dgvM_MakerZaiko.TabIndex = 3;
            this.dgvM_MakerZaiko.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvM_MakerZaiko_CellContentClick);
            this.dgvM_MakerZaiko.CellPainting += new System.Windows.Forms.DataGridViewCellPaintingEventHandler(this.dgvM_MakerZaiko_CellPainting);
            this.dgvM_MakerZaiko.ColumnWidthChanged += new System.Windows.Forms.DataGridViewColumnEventHandler(this.dgvM_MakerZaiko_ColumnWidthChanged);
            this.dgvM_MakerZaiko.EditingControlShowing += new System.Windows.Forms.DataGridViewEditingControlShowingEventHandler(this.dgvM_MakerZaiko_EditingControlShowing);
            this.dgvM_MakerZaiko.Paint += new System.Windows.Forms.PaintEventHandler(this.dgvM_MakerZaiko_Paint);
            this.dgvM_MakerZaiko.PreviewKeyDown += new System.Windows.Forms.PreviewKeyDownEventHandler(this.dgvM_MakerZaiko_PreviewKeyDown);
            // 
            // colNo
            // 
            this.colNo.DataPropertyName = "No";
            dataGridViewCellStyle19.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            this.colNo.DefaultCellStyle = dataGridViewCellStyle19;
            this.colNo.HeaderText = "No";
            this.colNo.Name = "colNo";
            this.colNo.ReadOnly = true;
            this.colNo.Width = 50;
            // 
            // Column1
            // 
            this.Column1.HeaderText = "Column1";
            this.Column1.Name = "Column1";
            this.Column1.Visible = false;
            // 
            // Column2
            // 
            this.Column2.HeaderText = "Column2";
            this.Column2.Name = "Column2";
            this.Column2.Visible = false;
            // 
            // colMakerCd
            // 
            this.colMakerCd.DataPropertyName = "MakerCD";
            this.colMakerCd.HeaderText = "メーカー(仕入先)";
            this.colMakerCd.MaxInputLength = 6;
            this.colMakerCd.MinimumWidth = 6;
            this.colMakerCd.Name = "colMakerCd";
            this.colMakerCd.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.colMakerCd.Width = 130;
            // 
            // btnMaker
            // 
            this.btnMaker.HeaderText = "";
            this.btnMaker.Name = "btnMaker";
            this.btnMaker.Text = "...";
            this.btnMaker.UseColumnTextForButtonValue = true;
            this.btnMaker.Width = 30;
            // 
            // colMakerName
            // 
            this.colMakerName.DataPropertyName = "MakerName";
            this.colMakerName.HeaderText = "";
            this.colMakerName.MaxInputLength = 25;
            this.colMakerName.MinimumWidth = 40;
            this.colMakerName.Name = "colMakerName";
            this.colMakerName.ReadOnly = true;
            this.colMakerName.Width = 250;
            // 
            // colBrandCD
            // 
            this.colBrandCD.DataPropertyName = "BrandCD";
            this.colBrandCD.HeaderText = "ブランド";
            this.colBrandCD.MaxInputLength = 3;
            this.colBrandCD.Name = "colBrandCD";
            this.colBrandCD.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.colBrandCD.Width = 90;
            // 
            // btnBrand
            // 
            this.btnBrand.HeaderText = "";
            this.btnBrand.Name = "btnBrand";
            this.btnBrand.Text = "...";
            this.btnBrand.UseColumnTextForButtonValue = true;
            this.btnBrand.Width = 30;
            // 
            // colBrandName
            // 
            this.colBrandName.DataPropertyName = "BrandName";
            this.colBrandName.HeaderText = "";
            this.colBrandName.MaxInputLength = 20;
            this.colBrandName.MinimumWidth = 40;
            this.colBrandName.Name = "colBrandName";
            this.colBrandName.ReadOnly = true;
            this.colBrandName.Width = 250;
            // 
            // colData
            // 
            this.colData.DataPropertyName = "DataSourseMakerCD";
            this.colData.HeaderText = "データ元(仕入先)";
            this.colData.MaxInputLength = 6;
            this.colData.MinimumWidth = 6;
            this.colData.Name = "colData";
            this.colData.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.colData.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.colData.Width = 130;
            // 
            // btnData
            // 
            this.btnData.HeaderText = "";
            this.btnData.Name = "btnData";
            this.btnData.Text = "...";
            this.btnData.UseColumnTextForButtonValue = true;
            this.btnData.Width = 30;
            // 
            // colDataName
            // 
            this.colDataName.DataPropertyName = "DataMakerName";
            this.colDataName.HeaderText = "";
            this.colDataName.MaxInputLength = 25;
            this.colDataName.MinimumWidth = 40;
            this.colDataName.Name = "colDataName";
            this.colDataName.ReadOnly = true;
            this.colDataName.Width = 250;
            // 
            // colPatternCD
            // 
            this.colPatternCD.DataPropertyName = "PatternCD";
            this.colPatternCD.HeaderText = "取込パターン";
            this.colPatternCD.MaxInputLength = 3;
            this.colPatternCD.Name = "colPatternCD";
            this.colPatternCD.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.colPatternCD.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.colPatternCD.Width = 90;
            // 
            // btnPattern
            // 
            this.btnPattern.HeaderText = "";
            this.btnPattern.Name = "btnPattern";
            this.btnPattern.Text = "...";
            this.btnPattern.UseColumnTextForButtonValue = true;
            this.btnPattern.Width = 30;
            // 
            // colPatternName
            // 
            this.colPatternName.DataPropertyName = "PatternName";
            this.colPatternName.HeaderText = "";
            this.colPatternName.MaxInputLength = 20;
            this.colPatternName.MinimumWidth = 40;
            this.colPatternName.Name = "colPatternName";
            this.colPatternName.ReadOnly = true;
            this.colPatternName.Width = 250;
            // 
            // colIrregularKBN
            // 
            this.colIrregularKBN.DataPropertyName = "IrregularKBN";
            this.colIrregularKBN.HeaderText = "例外";
            this.colIrregularKBN.Name = "colIrregularKBN";
            this.colIrregularKBN.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.colIrregularKBN.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.colIrregularKBN.Visible = false;
            // 
            // smS_Label7
            // 
            this.smS_Label7.BackColor = System.Drawing.Color.Black;
            this.smS_Label7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.smS_Label7.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label7.Location = new System.Drawing.Point(12, 65);
            this.smS_Label7.Name = "smS_Label7";
            this.smS_Label7.Size = new System.Drawing.Size(1650, 2);
            this.smS_Label7.TabIndex = 48;
            // 
            // smS_Label4
            // 
            this.smS_Label4.AutoSize = true;
            this.smS_Label4.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label4.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label4.Location = new System.Drawing.Point(587, 13);
            this.smS_Label4.Name = "smS_Label4";
            this.smS_Label4.Size = new System.Drawing.Size(110, 12);
            this.smS_Label4.TabIndex = 50;
            this.smS_Label4.Text = "データ元(仕入先)";
            // 
            // uc_DatasourceMaker
            // 
            this.uc_DatasourceMaker.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            this.uc_DatasourceMaker.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.uc_DatasourceMaker.Location = new System.Drawing.Point(700, 7);
            this.uc_DatasourceMaker.Margin = new System.Windows.Forms.Padding(0);
            this.uc_DatasourceMaker.Name = "uc_DatasourceMaker";
            this.uc_DatasourceMaker.Size = new System.Drawing.Size(465, 25);
            this.uc_DatasourceMaker.TabIndex = 49;
            this.uc_DatasourceMaker.UC_Code = "";
            this.uc_DatasourceMaker.UC_Code_Width = 50;
            this.uc_DatasourceMaker.UC_Exists = false;
            this.uc_DatasourceMaker.UC_Flag = 0;
            this.uc_DatasourceMaker.UC_IsRequired = false;
            this.uc_DatasourceMaker.UC_Name = "";
            this.uc_DatasourceMaker.UC_Name_Width = 330;
            this.uc_DatasourceMaker.UC_NameVisible = true;
            this.uc_DatasourceMaker.UC_SearchEnable = true;
            this.uc_DatasourceMaker.UC_Type = SMS.CustomControls.UC_Search.Type.Maker;
            this.uc_DatasourceMaker.Enter += new System.EventHandler(this.uc_DatasourceMaker_Enter);
            this.uc_DatasourceMaker.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uc_DatasourceMaker_KeyDown);
            this.uc_DatasourceMaker.Leave += new System.EventHandler(this.uc_Leave);
            // 
            // frmPSKS0102M
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1884, 951);
            this.KeyPreview = true;
            this.Location = new System.Drawing.Point(0, 0);
            this.Name = "frmPSKS0102M";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "PSKS0102M";
            this.Load += new System.EventHandler(this.frmPSKS0102M_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmPSKS0102M_KeyDown);
            this.panelBody.ResumeLayout(false);
            this.panelBody.PerformLayout();
            this.panelTitleLeft.ResumeLayout(false);
            this.panelTitle.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvM_MakerZaiko)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private CustomControls.SMS_Button button11;
        private CustomControls.SMS_Label smS_Label2;
        private CustomControls.SMS_Label smS_Label3;
        private CustomControls.UC_Search ucBrand;
        private CustomControls.UC_Search ucMaker;
        private SMS_GridView dgvM_MakerZaiko;
        private CustomControls.SMS_Label smS_Label7;
        private System.Windows.Forms.DataGridViewTextBoxColumn colNo;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column2;
        private System.Windows.Forms.DataGridViewTextBoxColumn colMakerCd;
        private System.Windows.Forms.DataGridViewButtonColumn btnMaker;
        private System.Windows.Forms.DataGridViewTextBoxColumn colMakerName;
        private System.Windows.Forms.DataGridViewTextBoxColumn colBrandCD;
        private System.Windows.Forms.DataGridViewButtonColumn btnBrand;
        private System.Windows.Forms.DataGridViewTextBoxColumn colBrandName;
        private System.Windows.Forms.DataGridViewTextBoxColumn colData;
        private System.Windows.Forms.DataGridViewButtonColumn btnData;
        private System.Windows.Forms.DataGridViewTextBoxColumn colDataName;
        private System.Windows.Forms.DataGridViewTextBoxColumn colPatternCD;
        private System.Windows.Forms.DataGridViewButtonColumn btnPattern;
        private System.Windows.Forms.DataGridViewTextBoxColumn colPatternName;
        private System.Windows.Forms.DataGridViewCheckBoxColumn colIrregularKBN;
        private CustomControls.SMS_Label smS_Label4;
        private CustomControls.UC_Search uc_DatasourceMaker;
    }
}