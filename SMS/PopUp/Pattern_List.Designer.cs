namespace SMS.PopUp
{
    partial class frmPattern_List
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmPattern_List));
            this.lblFormName = new System.Windows.Forms.Label();
            this.dgvMKZH = new SMS.SMS_GridView();
            this.colno = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colPatternCD = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colPatternName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btnF12 = new SMS.CustomControls.SMS_Button();
            this.btnF11 = new SMS.CustomControls.SMS_Button();
            this.btnF1 = new SMS.CustomControls.SMS_Button();
            this.txtPatternName = new SMS.CustomControls.SMS_TextBox();
            this.smS_Label1 = new SMS.CustomControls.SMS_Label();
            ((System.ComponentModel.ISupportInitialize)(this.dgvMKZH)).BeginInit();
            this.SuspendLayout();
            // 
            // lblFormName
            // 
            this.lblFormName.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(173)))), ((int)(((byte)(71)))));
            this.lblFormName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblFormName.Font = new System.Drawing.Font("MS Gothic", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFormName.Location = new System.Drawing.Point(45, 9);
            this.lblFormName.Name = "lblFormName";
            this.lblFormName.Size = new System.Drawing.Size(343, 22);
            this.lblFormName.TabIndex = 1;
            this.lblFormName.Text = "取込パターン検索";
            this.lblFormName.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // dgvMKZH
            // 
            this.dgvMKZH.AllowUserToAddRows = false;
            this.dgvMKZH.AllowUserToDeleteRows = false;
            this.dgvMKZH.AllowUserToResizeRows = false;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(221)))), ((int)(((byte)(235)))), ((int)(((byte)(247)))));
            this.dgvMKZH.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dgvMKZH.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            dataGridViewCellStyle2.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgvMKZH.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle2;
            this.dgvMKZH.ColumnHeadersHeight = 25;
            this.dgvMKZH.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colno,
            this.colPatternCD,
            this.colPatternName});
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("MS Gothic", 9F);
            dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgvMKZH.DefaultCellStyle = dataGridViewCellStyle3;
            this.dgvMKZH.EnableHeadersVisualStyles = false;
            this.dgvMKZH.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            this.dgvMKZH.IsCellValid = false;
            this.dgvMKZH.Location = new System.Drawing.Point(15, 68);
            this.dgvMKZH.Name = "dgvMKZH";
            this.dgvMKZH.RowHeadersVisible = false;
            this.dgvMKZH.Size = new System.Drawing.Size(420, 292);
            this.dgvMKZH.TabIndex = 1;
            this.dgvMKZH.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvMKZH_CellDoubleClick);
            // 
            // colno
            // 
            this.colno.DataPropertyName = "No";
            this.colno.HeaderText = "No";
            this.colno.Name = "colno";
            this.colno.Width = 50;
            // 
            // colPatternCD
            // 
            this.colPatternCD.DataPropertyName = "PatternCD";
            this.colPatternCD.FillWeight = 70F;
            this.colPatternCD.HeaderText = "パターンCD";
            this.colPatternCD.MinimumWidth = 3;
            this.colPatternCD.Name = "colPatternCD";
            this.colPatternCD.Width = 80;
            // 
            // colPatternName
            // 
            this.colPatternName.DataPropertyName = "PatternName";
            this.colPatternName.HeaderText = "パターン名";
            this.colPatternName.MinimumWidth = 20;
            this.colPatternName.Name = "colPatternName";
            this.colPatternName.Width = 270;
            // 
            // btnF12
            // 
            this.btnF12.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.btnF12.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnF12.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnF12.Location = new System.Drawing.Point(362, 372);
            this.btnF12.Name = "btnF12";
            this.btnF12.Size = new System.Drawing.Size(81, 30);
            this.btnF12.TabIndex = 4;
            this.btnF12.Text = "確定(F12)";
            this.btnF12.UseVisualStyleBackColor = false;
            this.btnF12.Click += new System.EventHandler(this.btnF12_Click);
            // 
            // btnF11
            // 
            this.btnF11.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.btnF11.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnF11.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnF11.Location = new System.Drawing.Point(275, 373);
            this.btnF11.Name = "btnF11";
            this.btnF11.Size = new System.Drawing.Size(81, 30);
            this.btnF11.TabIndex = 3;
            this.btnF11.Text = "表示(F11)";
            this.btnF11.UseVisualStyleBackColor = false;
            this.btnF11.Click += new System.EventHandler(this.btnF11_Click);
            // 
            // btnF1
            // 
            this.btnF1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.btnF1.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnF1.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnF1.Location = new System.Drawing.Point(10, 373);
            this.btnF1.Name = "btnF1";
            this.btnF1.Size = new System.Drawing.Size(81, 30);
            this.btnF1.TabIndex = 2;
            this.btnF1.Text = "戻る(F1)";
            this.btnF1.UseVisualStyleBackColor = false;
            this.btnF1.Click += new System.EventHandler(this.btnF1_Click);
            // 
            // txtPatternName
            // 
            this.txtPatternName.Decimalplace = ((byte)(0));
            this.txtPatternName.Location = new System.Drawing.Point(96, 38);
            this.txtPatternName.MaxLength = 20;
            this.txtPatternName.Name = "txtPatternName";
            this.txtPatternName.Size = new System.Drawing.Size(270, 19);
            this.txtPatternName.TabIndex = 0;
            this.txtPatternName.txt_Type = SMS.CustomControls.SMS_TextBox.Type.Hiragana;
            this.txtPatternName.UseMinus = false;
            this.txtPatternName.UseThousandSeparator = true;
            // 
            // smS_Label1
            // 
            this.smS_Label1.AutoSize = true;
            this.smS_Label1.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label1.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label1.Location = new System.Drawing.Point(19, 43);
            this.smS_Label1.Name = "smS_Label1";
            this.smS_Label1.Size = new System.Drawing.Size(70, 12);
            this.smS_Label1.TabIndex = 2;
            this.smS_Label1.Text = "パターン名";
            // 
            // frmPattern_List
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            this.ClientSize = new System.Drawing.Size(447, 409);
            this.Controls.Add(this.dgvMKZH);
            this.Controls.Add(this.btnF12);
            this.Controls.Add(this.btnF11);
            this.Controls.Add(this.btnF1);
            this.Controls.Add(this.txtPatternName);
            this.Controls.Add(this.smS_Label1);
            this.Controls.Add(this.lblFormName);
            this.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.KeyPreview = true;
            this.MaximizeBox = false;
            this.Name = "frmPattern_List";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "取込パターン検索画面";
            this.Load += new System.EventHandler(this.frmPattern_List_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmPattern_List_KeyDown);
            ((System.ComponentModel.ISupportInitialize)(this.dgvMKZH)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblFormName;
        private CustomControls.SMS_Label smS_Label1;
        private CustomControls.SMS_TextBox txtPatternName;
        private SMS_GridView dgvMKZH;
        private CustomControls.SMS_Button btnF1;
        private CustomControls.SMS_Button btnF11;
        private CustomControls.SMS_Button btnF12;
        private System.Windows.Forms.DataGridViewTextBoxColumn colno;
        private System.Windows.Forms.DataGridViewTextBoxColumn colPatternCD;
        private System.Windows.Forms.DataGridViewTextBoxColumn colPatternName;
    }
}