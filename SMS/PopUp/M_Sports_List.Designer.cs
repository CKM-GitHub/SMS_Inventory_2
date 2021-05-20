namespace SMS.PopUp
{
    partial class frmM_Sports_List
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmM_Sports_List));
            this.lblFormName = new System.Windows.Forms.Label();
            this.btnF12 = new SMS.CustomControls.SMS_Button();
            this.btnF11 = new SMS.CustomControls.SMS_Button();
            this.btnF1 = new SMS.CustomControls.SMS_Button();
            this.gvkyogi = new SMS.SMS_GridView();
            this.colNo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colSportsCD = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colSportsName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.txtvmsports = new SMS.CustomControls.SMS_TextBox();
            this.smS_Label1 = new SMS.CustomControls.SMS_Label();
            ((System.ComponentModel.ISupportInitialize)(this.gvkyogi)).BeginInit();
            this.SuspendLayout();
            // 
            // lblFormName
            // 
            this.lblFormName.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(173)))), ((int)(((byte)(71)))));
            this.lblFormName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblFormName.Font = new System.Drawing.Font("MS Gothic", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFormName.Location = new System.Drawing.Point(37, 8);
            this.lblFormName.Name = "lblFormName";
            this.lblFormName.Size = new System.Drawing.Size(343, 22);
            this.lblFormName.TabIndex = 4;
            this.lblFormName.Text = "競技検索画面";
            this.lblFormName.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btnF12
            // 
            this.btnF12.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.btnF12.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnF12.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnF12.Location = new System.Drawing.Point(323, 341);
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
            this.btnF11.Location = new System.Drawing.Point(242, 341);
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
            this.btnF1.Location = new System.Drawing.Point(13, 341);
            this.btnF1.Name = "btnF1";
            this.btnF1.Size = new System.Drawing.Size(75, 28);
            this.btnF1.TabIndex = 2;
            this.btnF1.Text = "戻る(F1)";
            this.btnF1.UseVisualStyleBackColor = false;
            this.btnF1.Click += new System.EventHandler(this.btnF1_Click);
            // 
            // gvkyogi
            // 
            this.gvkyogi.AllowUserToAddRows = false;
            this.gvkyogi.AllowUserToDeleteRows = false;
            this.gvkyogi.AllowUserToResizeRows = false;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(221)))), ((int)(((byte)(235)))), ((int)(((byte)(247)))));
            this.gvkyogi.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.gvkyogi.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            dataGridViewCellStyle2.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.gvkyogi.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle2;
            this.gvkyogi.ColumnHeadersHeight = 25;
            this.gvkyogi.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colNo,
            this.colSportsCD,
            this.colSportsName});
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("MS Gothic", 9F);
            dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.gvkyogi.DefaultCellStyle = dataGridViewCellStyle3;
            this.gvkyogi.EnableHeadersVisualStyles = false;
            this.gvkyogi.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            this.gvkyogi.IsCellValid = false;
            this.gvkyogi.Location = new System.Drawing.Point(13, 66);
            this.gvkyogi.Name = "gvkyogi";
            this.gvkyogi.RowHeadersVisible = false;
            this.gvkyogi.Size = new System.Drawing.Size(385, 268);
            this.gvkyogi.TabIndex = 5;
            this.gvkyogi.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.gvkyogi_CellDoubleClick);
            // 
            // colNo
            // 
            this.colNo.DataPropertyName = "No";
            this.colNo.HeaderText = "No";
            this.colNo.Name = "colNo";
            this.colNo.ReadOnly = true;
            this.colNo.Width = 50;
            // 
            // colSportsCD
            // 
            this.colSportsCD.DataPropertyName = "KyogiCD";
            this.colSportsCD.HeaderText = "競技CD";
            this.colSportsCD.MinimumWidth = 3;
            this.colSportsCD.Name = "colSportsCD";
            this.colSportsCD.ReadOnly = true;
            this.colSportsCD.Width = 60;
            // 
            // colSportsName
            // 
            this.colSportsName.DataPropertyName = "KyogiName";
            this.colSportsName.HeaderText = "競技名";
            this.colSportsName.MaxInputLength = 20;
            this.colSportsName.MinimumWidth = 100;
            this.colSportsName.Name = "colSportsName";
            this.colSportsName.ReadOnly = true;
            this.colSportsName.Width = 250;
            // 
            // txtvmsports
            // 
            this.txtvmsports.Decimalplace = ((byte)(0));
            this.txtvmsports.Location = new System.Drawing.Point(67, 38);
            this.txtvmsports.MaxLength = 20;
            this.txtvmsports.Name = "txtvmsports";
            this.txtvmsports.Size = new System.Drawing.Size(250, 19);
            this.txtvmsports.TabIndex = 0;
            this.txtvmsports.txt_Type = SMS.CustomControls.SMS_TextBox.Type.Hiragana;
            this.txtvmsports.UseMinus = false;
            this.txtvmsports.UseThousandSeparator = true;
            // 
            // smS_Label1
            // 
            this.smS_Label1.AutoSize = true;
            this.smS_Label1.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label1.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold);
            this.smS_Label1.Location = new System.Drawing.Point(12, 42);
            this.smS_Label1.Name = "smS_Label1";
            this.smS_Label1.Size = new System.Drawing.Size(44, 12);
            this.smS_Label1.TabIndex = 6;
            this.smS_Label1.Text = "競技名";
            // 
            // frmM_Sports_List
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            this.ClientSize = new System.Drawing.Size(416, 377);
            this.Controls.Add(this.btnF12);
            this.Controls.Add(this.btnF11);
            this.Controls.Add(this.btnF1);
            this.Controls.Add(this.gvkyogi);
            this.Controls.Add(this.txtvmsports);
            this.Controls.Add(this.smS_Label1);
            this.Controls.Add(this.lblFormName);
            this.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "frmM_Sports_List";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "競技検索画面";
            this.Load += new System.EventHandler(this.M_Sports_List_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.M_Sports_List_KeyDown);
            ((System.ComponentModel.ISupportInitialize)(this.gvkyogi)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblFormName;
        private CustomControls.SMS_TextBox txtvmsports;
        private CustomControls.SMS_Label smS_Label1;
        private SMS_GridView gvkyogi;
        private System.Windows.Forms.DataGridViewTextBoxColumn colNo;
        private System.Windows.Forms.DataGridViewTextBoxColumn colSportsCD;
        private System.Windows.Forms.DataGridViewTextBoxColumn colSportsName;
        private CustomControls.SMS_Button btnF12;
        private CustomControls.SMS_Button btnF11;
        private CustomControls.SMS_Button btnF1;
    }
}