namespace CKM_Controls
{
    partial class ComboBoxTest
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
            this.components = new System.ComponentModel.Container();
            this.testBoxTime1 = new CKM_Controls.TestBoxTime(this.components);
            this.ckM_ComboBox1 = new CKM_Controls.CKM_ComboBox(this.components);
            this.SuspendLayout();
            // 
            // testBoxTime1
            // 
            this.testBoxTime1.CKM_Reqired = CKM_Controls.TestBoxTime.Type.Time;
            this.testBoxTime1.CKM_Type = CKM_Controls.TestBoxTime.Type.Time;
            this.testBoxTime1.Location = new System.Drawing.Point(42, 65);
            this.testBoxTime1.MaxLength = 6;
            this.testBoxTime1.Name = "testBoxTime1";
            this.testBoxTime1.Size = new System.Drawing.Size(100, 20);
            this.testBoxTime1.TabIndex = 1;
            // 
            // ckM_ComboBox1
            // 
            this.ckM_ComboBox1.BackColor = System.Drawing.Color.White;
            this.ckM_ComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.ckM_ComboBox1.FormattingEnabled = true;
            this.ckM_ComboBox1.Location = new System.Drawing.Point(42, 21);
            this.ckM_ComboBox1.Name = "ckM_ComboBox1";
            this.ckM_ComboBox1.Size = new System.Drawing.Size(121, 21);
            this.ckM_ComboBox1.TabIndex = 0;
            // 
            // ComboBoxTest
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.Controls.Add(this.testBoxTime1);
            this.Controls.Add(this.ckM_ComboBox1);
            this.Name = "ComboBoxTest";
            this.Text = "ComboBoxTest";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private CKM_ComboBox ckM_ComboBox1;
        private TestBoxTime testBoxTime1;
    }
}