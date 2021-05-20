namespace SMS.Order
{
    partial class F0008
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
            this.btnUnCheckAll = new SMS.CustomControls.SMS_Button();
            this.btnCheckAll = new SMS.CustomControls.SMS_Button();
            this.dgv1 = new SMS.SMS_GridView();
            this.colNo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column2 = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.colSupplier = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colOrderCount = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colCaptureNo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colErrorNo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.smS_Button1 = new SMS.CustomControls.SMS_Button();
            this.smS_Label2 = new SMS.CustomControls.SMS_Label();
            this.dateTextBox1 = new SMS.CustomControls.DateTextBox();
            this.panelBody.SuspendLayout();
            this.panelTitleLeft.SuspendLayout();
            this.panelTitle.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).BeginInit();
            this.SuspendLayout();
            // 
            // panelBody
            // 
            this.panelBody.Controls.Add(this.dateTextBox1);
            this.panelBody.Controls.Add(this.btnUnCheckAll);
            this.panelBody.Controls.Add(this.btnCheckAll);
            this.panelBody.Controls.Add(this.dgv1);
            this.panelBody.Controls.Add(this.smS_Button1);
            this.panelBody.Controls.Add(this.smS_Label2);
            this.panelBody.TabIndex = 0;
            // 
            // btnUnCheckAll
            // 
            this.btnUnCheckAll.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.btnUnCheckAll.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnUnCheckAll.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnUnCheckAll.Location = new System.Drawing.Point(721, 798);
            this.btnUnCheckAll.Name = "btnUnCheckAll";
            this.btnUnCheckAll.Size = new System.Drawing.Size(109, 28);
            this.btnUnCheckAll.TabIndex = 11;
            this.btnUnCheckAll.Text = "全解除";
            this.btnUnCheckAll.UseVisualStyleBackColor = false;
            this.btnUnCheckAll.Click += new System.EventHandler(this.btnUnCheckAll_Click);
            // 
            // btnCheckAll
            // 
            this.btnCheckAll.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.btnCheckAll.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnCheckAll.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnCheckAll.Location = new System.Drawing.Point(613, 798);
            this.btnCheckAll.Name = "btnCheckAll";
            this.btnCheckAll.Size = new System.Drawing.Size(109, 28);
            this.btnCheckAll.TabIndex = 10;
            this.btnCheckAll.Text = "全選択";
            this.btnCheckAll.UseVisualStyleBackColor = false;
            this.btnCheckAll.Click += new System.EventHandler(this.btnCheckAll_Click);
            // 
            // dgv1
            // 
            this.dgv1.AllowUserToAddRows = false;
            this.dgv1.AllowUserToDeleteRows = false;
            this.dgv1.AllowUserToResizeRows = false;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(221)))), ((int)(((byte)(235)))), ((int)(((byte)(247)))));
            this.dgv1.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dgv1.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            dataGridViewCellStyle2.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgv1.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle2;
            this.dgv1.ColumnHeadersHeight = 25;
            this.dgv1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colNo,
            this.Column2,
            this.colSupplier,
            this.colOrderCount,
            this.colCaptureNo,
            this.colErrorNo});
            dataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle7.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle7.Font = new System.Drawing.Font("MS Gothic", 9F);
            dataGridViewCellStyle7.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle7.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle7.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgv1.DefaultCellStyle = dataGridViewCellStyle7;
            this.dgv1.EnableHeadersVisualStyles = false;
            this.dgv1.GridColor = System.Drawing.Color.DimGray;
            this.dgv1.IsCellValid = false;
            this.dgv1.Location = new System.Drawing.Point(178, 92);
            this.dgv1.Name = "dgv1";
            this.dgv1.RowHeadersVisible = false;
            this.dgv1.Size = new System.Drawing.Size(650, 700);
            this.dgv1.TabIndex = 2;
            // 
            // colNo
            // 
            this.colNo.DataPropertyName = "No";
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            this.colNo.DefaultCellStyle = dataGridViewCellStyle3;
            this.colNo.HeaderText = "No";
            this.colNo.Name = "colNo";
            this.colNo.Width = 30;
            // 
            // Column2
            // 
            this.Column2.HeaderText = "";
            this.Column2.Name = "Column2";
            this.Column2.Width = 50;
            // 
            // colSupplier
            // 
            this.colSupplier.DataPropertyName = "Supplier";
            this.colSupplier.HeaderText = "仕入先";
            this.colSupplier.Name = "colSupplier";
            this.colSupplier.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.colSupplier.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.colSupplier.Width = 350;
            // 
            // colOrderCount
            // 
            this.colOrderCount.DataPropertyName = "OrderCount";
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            this.colOrderCount.DefaultCellStyle = dataGridViewCellStyle4;
            this.colOrderCount.HeaderText = "発注数";
            this.colOrderCount.Name = "colOrderCount";
            this.colOrderCount.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.colOrderCount.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.colOrderCount.Width = 60;
            // 
            // colCaptureNo
            // 
            this.colCaptureNo.DataPropertyName = "CaptureNo";
            dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            this.colCaptureNo.DefaultCellStyle = dataGridViewCellStyle5;
            this.colCaptureNo.HeaderText = "取込数";
            this.colCaptureNo.Name = "colCaptureNo";
            this.colCaptureNo.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.colCaptureNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.colCaptureNo.Width = 60;
            // 
            // colErrorNo
            // 
            this.colErrorNo.DataPropertyName = "ErrorNo";
            dataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            this.colErrorNo.DefaultCellStyle = dataGridViewCellStyle6;
            this.colErrorNo.HeaderText = "エラー数";
            this.colErrorNo.Name = "colErrorNo";
            this.colErrorNo.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.colErrorNo.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            this.colErrorNo.Width = 70;
            // 
            // smS_Button1
            // 
            this.smS_Button1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.smS_Button1.Cursor = System.Windows.Forms.Cursors.Hand;
            this.smS_Button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.smS_Button1.Location = new System.Drawing.Point(720, 59);
            this.smS_Button1.Name = "smS_Button1";
            this.smS_Button1.Size = new System.Drawing.Size(109, 28);
            this.smS_Button1.TabIndex = 1;
            this.smS_Button1.Text = "再表示(F11)";
            this.smS_Button1.UseVisualStyleBackColor = false;
            // 
            // smS_Label2
            // 
            this.smS_Label2.AutoSize = true;
            this.smS_Label2.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label2.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label2.Location = new System.Drawing.Point(577, 70);
            this.smS_Label2.Name = "smS_Label2";
            this.smS_Label2.Size = new System.Drawing.Size(44, 12);
            this.smS_Label2.TabIndex = 6;
            this.smS_Label2.Text = "取込日";
            // 
            // dateTextBox1
            // 
            this.dateTextBox1.Location = new System.Drawing.Point(624, 66);
            this.dateTextBox1.Mask = "0000/00/00";
            this.dateTextBox1.MaskFormat = "yyyy/MM/dd";
            this.dateTextBox1.Name = "dateTextBox1";
            this.dateTextBox1.PromptChar = ' ';
            this.dateTextBox1.Size = new System.Drawing.Size(88, 19);
            this.dateTextBox1.TabIndex = 0;
            // 
            // F0008
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1884, 951);
            this.F10_Text = "リスト印刷(F10)";
            this.F12_Text = "データ出力(F12)";
            this.F5_Text = "ｷｬﾝｾﾙ(F5)";
            this.KeyPreview = true;
            this.Location = new System.Drawing.Point(0, 0);
            this.Name = "F0008";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "F0008";
            this.Load += new System.EventHandler(this.F0008_Load);
            this.panelBody.ResumeLayout(false);
            this.panelBody.PerformLayout();
            this.panelTitleLeft.ResumeLayout(false);
            this.panelTitle.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private CustomControls.SMS_Button btnUnCheckAll;
        private CustomControls.SMS_Button btnCheckAll;
        private SMS_GridView dgv1;
        private CustomControls.SMS_Button smS_Button1;
        private CustomControls.SMS_Label smS_Label2;
        private CustomControls.DateTextBox dateTextBox1;
        private System.Windows.Forms.DataGridViewTextBoxColumn colNo;
        private System.Windows.Forms.DataGridViewCheckBoxColumn Column2;
        private System.Windows.Forms.DataGridViewTextBoxColumn colSupplier;
        private System.Windows.Forms.DataGridViewTextBoxColumn colOrderCount;
        private System.Windows.Forms.DataGridViewTextBoxColumn colCaptureNo;
        private System.Windows.Forms.DataGridViewTextBoxColumn colErrorNo;
    }
}