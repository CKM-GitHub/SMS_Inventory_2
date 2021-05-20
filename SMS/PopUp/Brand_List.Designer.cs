namespace SMS.PopUp
{
    partial class frmBrand_List
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmBrand_List));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            this.lblFormName = new System.Windows.Forms.Label();
            this.btnF12 = new SMS.CustomControls.SMS_Button();
            this.btnF11 = new SMS.CustomControls.SMS_Button();
            this.btnF1 = new SMS.CustomControls.SMS_Button();
            this.dgvBrand = new SMS.SMS_GridView();
            this.colNo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colBrandCD = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colBrandName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.txtBrandName = new SMS.CustomControls.SMS_TextBox();
            this.smS_Label1 = new SMS.CustomControls.SMS_Label();
            ((System.ComponentModel.ISupportInitialize)(this.dgvBrand)).BeginInit();
            this.SuspendLayout();
            // 
            // lblFormName
            // 
            this.lblFormName.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(112)))), ((int)(((byte)(173)))), ((int)(((byte)(71)))));
            this.lblFormName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            resources.ApplyResources(this.lblFormName, "lblFormName");
            this.lblFormName.Name = "lblFormName";
            // 
            // btnF12
            // 
            this.btnF12.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.btnF12.Cursor = System.Windows.Forms.Cursors.Hand;
            resources.ApplyResources(this.btnF12, "btnF12");
            this.btnF12.Name = "btnF12";
            this.btnF12.UseVisualStyleBackColor = false;
            this.btnF12.Click += new System.EventHandler(this.btnF12_Click);
            // 
            // btnF11
            // 
            this.btnF11.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.btnF11.Cursor = System.Windows.Forms.Cursors.Hand;
            resources.ApplyResources(this.btnF11, "btnF11");
            this.btnF11.Name = "btnF11";
            this.btnF11.UseVisualStyleBackColor = false;
            this.btnF11.Click += new System.EventHandler(this.btnF11_Click);
            // 
            // btnF1
            // 
            this.btnF1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            this.btnF1.Cursor = System.Windows.Forms.Cursors.Hand;
            resources.ApplyResources(this.btnF1, "btnF1");
            this.btnF1.Name = "btnF1";
            this.btnF1.UseVisualStyleBackColor = false;
            this.btnF1.Click += new System.EventHandler(this.btnF1_Click);
            // 
            // dgvBrand
            // 
            this.dgvBrand.AllowUserToAddRows = false;
            this.dgvBrand.AllowUserToDeleteRows = false;
            this.dgvBrand.AllowUserToResizeRows = false;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(221)))), ((int)(((byte)(235)))), ((int)(((byte)(247)))));
            this.dgvBrand.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dgvBrand.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(191)))), ((int)(((byte)(191)))), ((int)(((byte)(191)))));
            dataGridViewCellStyle2.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgvBrand.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle2;
            resources.ApplyResources(this.dgvBrand, "dgvBrand");
            this.dgvBrand.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colNo,
            this.colBrandCD,
            this.colBrandName});
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("MS Gothic", 9F);
            dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgvBrand.DefaultCellStyle = dataGridViewCellStyle3;
            this.dgvBrand.EnableHeadersVisualStyles = false;
            this.dgvBrand.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            this.dgvBrand.IsCellValid = false;
            this.dgvBrand.Name = "dgvBrand";
            this.dgvBrand.RowHeadersVisible = false;
            this.dgvBrand.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvBrand_CellDoubleClick);
            // 
            // colNo
            // 
            this.colNo.DataPropertyName = "No";
            resources.ApplyResources(this.colNo, "colNo");
            this.colNo.Name = "colNo";
            // 
            // colBrandCD
            // 
            this.colBrandCD.DataPropertyName = "nc_brand";
            resources.ApplyResources(this.colBrandCD, "colBrandCD");
            this.colBrandCD.Name = "colBrandCD";
            // 
            // colBrandName
            // 
            this.colBrandName.DataPropertyName = "vm_brand";
            resources.ApplyResources(this.colBrandName, "colBrandName");
            this.colBrandName.MaxInputLength = 20;
            this.colBrandName.Name = "colBrandName";
            // 
            // txtBrandName
            // 
            this.txtBrandName.Decimalplace = ((byte)(0));
            resources.ApplyResources(this.txtBrandName, "txtBrandName");
            this.txtBrandName.Name = "txtBrandName";
            this.txtBrandName.txt_Type = SMS.CustomControls.SMS_TextBox.Type.Hiragana;
            this.txtBrandName.UseMinus = false;
            this.txtBrandName.UseThousandSeparator = true;
            // 
            // smS_Label1
            // 
            resources.ApplyResources(this.smS_Label1, "smS_Label1");
            this.smS_Label1.BackColor = System.Drawing.Color.Transparent;
            this.smS_Label1.Name = "smS_Label1";
            // 
            // frmBrand_List
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(198)))), ((int)(((byte)(224)))), ((int)(((byte)(180)))));
            this.Controls.Add(this.btnF12);
            this.Controls.Add(this.btnF11);
            this.Controls.Add(this.btnF1);
            this.Controls.Add(this.dgvBrand);
            this.Controls.Add(this.txtBrandName);
            this.Controls.Add(this.smS_Label1);
            this.Controls.Add(this.lblFormName);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.KeyPreview = true;
            this.MaximizeBox = false;
            this.Name = "frmBrand_List";
            this.Load += new System.EventHandler(this.frmBrand_List_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmBrand_List_KeyDown);
            ((System.ComponentModel.ISupportInitialize)(this.dgvBrand)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblFormName;
        private CustomControls.SMS_Label smS_Label1;
        private CustomControls.SMS_TextBox txtBrandName;
        private SMS_GridView dgvBrand;
        private CustomControls.SMS_Button btnF1;
        private CustomControls.SMS_Button btnF11;
        private CustomControls.SMS_Button btnF12;
        private System.Windows.Forms.DataGridViewTextBoxColumn colNo;
        private System.Windows.Forms.DataGridViewTextBoxColumn colBrandCD;
        private System.Windows.Forms.DataGridViewTextBoxColumn colBrandName;
    }
}