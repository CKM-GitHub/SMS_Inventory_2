namespace SMS.CustomControls
{
    partial class ucSearchCombo
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.smS_ComboBox1 = new SMS.CustomControls.SMS_ComboBox();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.SuspendLayout();
            // 
            // smS_ComboBox1
            // 
            this.smS_ComboBox1.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.smS_ComboBox1.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.smS_ComboBox1.Cbo_Type = SMS.CustomControls.SMS_ComboBox.cbo_Type.m_nendo;
            this.smS_ComboBox1.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.smS_ComboBox1.FormattingEnabled = true;
            this.smS_ComboBox1.Location = new System.Drawing.Point(0, 0);
            this.smS_ComboBox1.Name = "smS_ComboBox1";
            this.smS_ComboBox1.Size = new System.Drawing.Size(196, 21);
            this.smS_ComboBox1.TabIndex = 0;
            // 
            // listBox1
            // 
            this.listBox1.FormattingEnabled = true;
            this.listBox1.Location = new System.Drawing.Point(0, 22);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(196, 43);
            this.listBox1.TabIndex = 1;
            this.listBox1.Visible = false;
            // 
            // ucSearchCombo
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Transparent;
            this.Controls.Add(this.listBox1);
            this.Controls.Add(this.smS_ComboBox1);
            this.Name = "ucSearchCombo";
            this.Size = new System.Drawing.Size(314, 66);
            this.ResumeLayout(false);

        }

        #endregion

        private SMS_ComboBox smS_ComboBox1;
        private System.Windows.Forms.ListBox listBox1;
    }
}
