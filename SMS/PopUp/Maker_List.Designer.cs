namespace SMS.PopUp
{
    partial class frmMaker_List
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmMaker_List));
            this.lblFormName = new System.Windows.Forms.Label();
            this.smS_Label1 = new SMS.CustomControls.SMS_Label();
            this.txtMakerName = new SMS.CustomControls.SMS_TextBox();
            this.dgvMakerSearch = new SMS.SMS_GridView();
            this.colNo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colMakerCD = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colMakerName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btnF1 = new SMS.CustomControls.SMS_Button();
            this.btnF11 = new SMS.CustomControls.SMS_Button();
            this.btnF12 = new SMS.CustomControls.SMS_Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgvMakerSearch)).BeginInit();
            this.SuspendLayout();
            // 
            // lblFormName
            // 
            this.lblFormName.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(173)))), ((int)(((byte)(71)))));
            this.lblFormName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblFormName.Font = new System.Drawing.Font("MS Gothic", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFormName.Location = new System.Drawing.Point(88, 9);
            this.lblFormName.Name = "lblFormName";
            this.lblFormName.Size = new System.Drawing.Size(343, 22);
            this.lblFormName.TabIndex = 2;
            this.lblFormName.Text = "仕入先検索";
            this.lblFormName.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // smS_Label1
            // 
            this.smS_Label1.AutoSize = true;
            this.smS_Label1.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label1.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label1.Location = new System.Drawing.Point(21, 49);
            this.smS_Label1.Name = "smS_Label1";
            this.smS_Label1.Size = new System.Drawing.Size(44, 12);
            this.smS_Label1.TabIndex = 3;
            this.smS_Label1.Text = "仕入先";
            // 
            // txtMakerName
            // 
            this.txtMakerName.Decimalplace = ((byte)(0));
            this.txtMakerName.Location = new System.Drawing.Point(71, 46);
            this.txtMakerName.MaxLength = 20;
            this.txtMakerName.Name = "txtMakerName";
            this.txtMakerName.Size = new System.Drawing.Size(360, 19);
            this.txtMakerName.TabIndex = 0;
            this.txtMakerName.txt_Type = SMS.CustomControls.SMS_TextBox.Type.Hiragana;
            this.txtMakerName.UseMinus = false;
            this.txtMakerName.UseThousandSeparator = true;
            // 
            // dgvMakerSearch
            // 
            this.dgvMakerSearch.AllowUserToAddRows = false;
            this.dgvMakerSearch.AllowUserToDeleteRows = false;
            this.dgvMakerSearch.AllowUserToResizeRows = false;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(221)))), ((int)(((byte)(235)))), ((int)(((byte)(247)))));
            this.dgvMakerSearch.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dgvMakerSearch.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            dataGridViewCellStyle2.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgvMakerSearch.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle2;
            this.dgvMakerSearch.ColumnHeadersHeight = 25;
            this.dgvMakerSearch.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colNo,
            this.colMakerCD,
            this.colMakerName});
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("MS Gothic", 9F);
            dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgvMakerSearch.DefaultCellStyle = dataGridViewCellStyle3;
            this.dgvMakerSearch.EnableHeadersVisualStyles = false;
            this.dgvMakerSearch.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            this.dgvMakerSearch.IsCellValid = false;
            this.dgvMakerSearch.Location = new System.Drawing.Point(12, 81);
            this.dgvMakerSearch.Name = "dgvMakerSearch";
            this.dgvMakerSearch.RowHeadersVisible = false;
            this.dgvMakerSearch.Size = new System.Drawing.Size(495, 287);
            this.dgvMakerSearch.TabIndex = 1;
            this.dgvMakerSearch.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvMakerSearch_CellDoubleClick);
            // 
            // colNo
            // 
            this.colNo.DataPropertyName = "No";
            this.colNo.HeaderText = "No";
            this.colNo.Name = "colNo";
            this.colNo.Width = 50;
            // 
            // colMakerCD
            // 
            this.colMakerCD.DataPropertyName = "VC_SHIIRESAKI";
            this.colMakerCD.HeaderText = "仕入先CD";
            this.colMakerCD.MinimumWidth = 3;
            this.colMakerCD.Name = "colMakerCD";
            this.colMakerCD.Width = 65;
            // 
            // colMakerName
            // 
            this.colMakerName.DataPropertyName = "VM_SHIIRESAKI";
            this.colMakerName.HeaderText = "仕入先名";
            this.colMakerName.MaxInputLength = 20;
            this.colMakerName.MinimumWidth = 20;
            this.colMakerName.Name = "colMakerName";
            this.colMakerName.Width = 360;
            // 
            // btnF1
            // 
            this.btnF1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.btnF1.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnF1.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnF1.Location = new System.Drawing.Point(12, 380);
            this.btnF1.Name = "btnF1";
            this.btnF1.Size = new System.Drawing.Size(81, 30);
            this.btnF1.TabIndex = 2;
            this.btnF1.Text = "戻る(F1)";
            this.btnF1.UseVisualStyleBackColor = false;
            this.btnF1.Click += new System.EventHandler(this.btnF1_Click);
            // 
            // btnF11
            // 
            this.btnF11.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.btnF11.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnF11.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnF11.Location = new System.Drawing.Point(337, 380);
            this.btnF11.Name = "btnF11";
            this.btnF11.Size = new System.Drawing.Size(81, 30);
            this.btnF11.TabIndex = 3;
            this.btnF11.Text = "表示(F11)";
            this.btnF11.UseVisualStyleBackColor = false;
            this.btnF11.Click += new System.EventHandler(this.btnF11_Click);
            // 
            // btnF12
            // 
            this.btnF12.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.btnF12.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnF12.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnF12.Location = new System.Drawing.Point(422, 380);
            this.btnF12.Name = "btnF12";
            this.btnF12.Size = new System.Drawing.Size(81, 30);
            this.btnF12.TabIndex = 4;
            this.btnF12.Text = "確定(F12)";
            this.btnF12.UseVisualStyleBackColor = false;
            this.btnF12.Click += new System.EventHandler(this.btnF12_Click);
            // 
            // frmMaker_List
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            this.ClientSize = new System.Drawing.Size(515, 416);
            this.Controls.Add(this.btnF12);
            this.Controls.Add(this.btnF11);
            this.Controls.Add(this.btnF1);
            this.Controls.Add(this.dgvMakerSearch);
            this.Controls.Add(this.txtMakerName);
            this.Controls.Add(this.smS_Label1);
            this.Controls.Add(this.lblFormName);
            this.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.KeyPreview = true;
            this.MaximizeBox = false;
            this.Name = "frmMaker_List";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "仕入先検索画面";
            this.Load += new System.EventHandler(this.frmMaker_List_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmMaker_List_KeyDown);
            ((System.ComponentModel.ISupportInitialize)(this.dgvMakerSearch)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblFormName;
        private CustomControls.SMS_Label smS_Label1;
        private CustomControls.SMS_TextBox txtMakerName;
        private SMS_GridView dgvMakerSearch;
        private CustomControls.SMS_Button btnF1;
        private CustomControls.SMS_Button btnF11;
        private CustomControls.SMS_Button btnF12;
        private System.Windows.Forms.DataGridViewTextBoxColumn colNo;
        private System.Windows.Forms.DataGridViewTextBoxColumn colMakerCD;
        private System.Windows.Forms.DataGridViewTextBoxColumn colMakerName;
    }
}