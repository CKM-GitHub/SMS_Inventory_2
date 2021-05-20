namespace SMS.PopUp
{
    partial class frmM_Bunrui_List
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmM_Bunrui_List));
            this.btnF12 = new SMS.CustomControls.SMS_Button();
            this.btnF11 = new SMS.CustomControls.SMS_Button();
            this.btnF1 = new SMS.CustomControls.SMS_Button();
            this.dgvBunrui = new SMS.SMS_GridView();
            this.colNo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colCategoryCD = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colCategoryName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.txtvmbunrui = new SMS.CustomControls.SMS_TextBox();
            this.smS_Label1 = new SMS.CustomControls.SMS_Label();
            this.lblFormName = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dgvBunrui)).BeginInit();
            this.SuspendLayout();
            // 
            // btnF12
            // 
            this.btnF12.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.btnF12.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnF12.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnF12.Location = new System.Drawing.Point(364, 344);
            this.btnF12.Name = "btnF12";
            this.btnF12.Size = new System.Drawing.Size(75, 28);
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
            this.btnF11.Location = new System.Drawing.Point(283, 344);
            this.btnF11.Name = "btnF11";
            this.btnF11.Size = new System.Drawing.Size(75, 28);
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
            this.btnF1.Location = new System.Drawing.Point(14, 344);
            this.btnF1.Name = "btnF1";
            this.btnF1.Size = new System.Drawing.Size(75, 28);
            this.btnF1.TabIndex = 2;
            this.btnF1.Text = "戻る(F1)";
            this.btnF1.UseVisualStyleBackColor = false;
            this.btnF1.Click += new System.EventHandler(this.btnF1_Click);
            // 
            // dgvBunrui
            // 
            this.dgvBunrui.AllowUserToAddRows = false;
            this.dgvBunrui.AllowUserToDeleteRows = false;
            this.dgvBunrui.AllowUserToResizeRows = false;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(221)))), ((int)(((byte)(235)))), ((int)(((byte)(247)))));
            this.dgvBunrui.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dgvBunrui.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            dataGridViewCellStyle2.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgvBunrui.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle2;
            this.dgvBunrui.ColumnHeadersHeight = 25;
            this.dgvBunrui.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colNo,
            this.colCategoryCD,
            this.colCategoryName});
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("MS Gothic", 9F);
            dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgvBunrui.DefaultCellStyle = dataGridViewCellStyle3;
            this.dgvBunrui.EnableHeadersVisualStyles = false;
            this.dgvBunrui.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            this.dgvBunrui.IsCellValid = false;
            this.dgvBunrui.Location = new System.Drawing.Point(14, 69);
            this.dgvBunrui.Name = "dgvBunrui";
            this.dgvBunrui.RowHeadersVisible = false;
            this.dgvBunrui.Size = new System.Drawing.Size(425, 268);
            this.dgvBunrui.TabIndex = 1;
            this.dgvBunrui.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvBunrui_CellDoubleClick);
            // 
            // colNo
            // 
            this.colNo.DataPropertyName = "No";
            this.colNo.HeaderText = "No";
            this.colNo.Name = "colNo";
            this.colNo.ReadOnly = true;
            this.colNo.Width = 50;
            // 
            // colCategoryCD
            // 
            this.colCategoryCD.DataPropertyName = "Katagori_CD";
            this.colCategoryCD.HeaderText = "カテゴリーCD";
            this.colCategoryCD.MinimumWidth = 3;
            this.colCategoryCD.Name = "colCategoryCD";
            this.colCategoryCD.ReadOnly = true;
            // 
            // colCategoryName
            // 
            this.colCategoryName.DataPropertyName = "Katagori_Name";
            this.colCategoryName.HeaderText = "カテゴリー名";
            this.colCategoryName.MaxInputLength = 20;
            this.colCategoryName.MinimumWidth = 100;
            this.colCategoryName.Name = "colCategoryName";
            this.colCategoryName.ReadOnly = true;
            this.colCategoryName.Width = 250;
            // 
            // txtvmbunrui
            // 
            this.txtvmbunrui.Decimalplace = ((byte)(0));
            this.txtvmbunrui.Location = new System.Drawing.Point(88, 39);
            this.txtvmbunrui.MaxLength = 20;
            this.txtvmbunrui.Name = "txtvmbunrui";
            this.txtvmbunrui.Size = new System.Drawing.Size(250, 19);
            this.txtvmbunrui.TabIndex = 0;
            this.txtvmbunrui.txt_Type = SMS.CustomControls.SMS_TextBox.Type.Hiragana;
            this.txtvmbunrui.UseMinus = false;
            this.txtvmbunrui.UseThousandSeparator = true;
            // 
            // smS_Label1
            // 
            this.smS_Label1.AutoSize = true;
            this.smS_Label1.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label1.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label1.Location = new System.Drawing.Point(13, 42);
            this.smS_Label1.Name = "smS_Label1";
            this.smS_Label1.Size = new System.Drawing.Size(70, 12);
            this.smS_Label1.TabIndex = 14;
            this.smS_Label1.Text = "カテゴリー";
            // 
            // lblFormName
            // 
            this.lblFormName.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(173)))), ((int)(((byte)(71)))));
            this.lblFormName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblFormName.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFormName.Location = new System.Drawing.Point(52, 8);
            this.lblFormName.Name = "lblFormName";
            this.lblFormName.Size = new System.Drawing.Size(343, 22);
            this.lblFormName.TabIndex = 13;
            this.lblFormName.Text = "カテゴリー検索画面";
            this.lblFormName.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // frmM_Bunrui_List
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            this.ClientSize = new System.Drawing.Size(449, 383);
            this.Controls.Add(this.btnF12);
            this.Controls.Add(this.btnF11);
            this.Controls.Add(this.btnF1);
            this.Controls.Add(this.dgvBunrui);
            this.Controls.Add(this.txtvmbunrui);
            this.Controls.Add(this.smS_Label1);
            this.Controls.Add(this.lblFormName);
            this.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "frmM_Bunrui_List";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "M_Bunrui_List";
            this.Load += new System.EventHandler(this.frmM_Bunrui_List_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmM_Bunrui_List_KeyDown);
            ((System.ComponentModel.ISupportInitialize)(this.dgvBunrui)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private CustomControls.SMS_Button btnF12;
        private CustomControls.SMS_Button btnF11;
        private CustomControls.SMS_Button btnF1;
        private SMS_GridView dgvBunrui;
        private CustomControls.SMS_TextBox txtvmbunrui;
        private CustomControls.SMS_Label smS_Label1;
        private System.Windows.Forms.Label lblFormName;
        private System.Windows.Forms.DataGridViewTextBoxColumn colNo;
        private System.Windows.Forms.DataGridViewTextBoxColumn colCategoryCD;
        private System.Windows.Forms.DataGridViewTextBoxColumn colCategoryName;
    }
}