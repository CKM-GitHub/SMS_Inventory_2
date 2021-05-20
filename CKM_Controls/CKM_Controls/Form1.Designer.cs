namespace CKM_Controls
{
    partial class Form1
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
            this.ckM_TextBox1 = new CKM_Controls.CKM_TextBox();
            this.SuspendLayout();
            // 
            // ckM_TextBox1
            // 
            this.ckM_TextBox1.CKM_Reqired = true;
            this.ckM_TextBox1.CKM_Type = CKM_Controls.CKM_TextBox.Type.Normal;
            this.ckM_TextBox1.Location = new System.Drawing.Point(42, 44);
            this.ckM_TextBox1.Name = "ckM_TextBox1";
            this.ckM_TextBox1.Size = new System.Drawing.Size(100, 20);
            this.ckM_TextBox1.TabIndex = 0;
            this.ckM_TextBox1.UseThousandSeparator = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.Controls.Add(this.ckM_TextBox1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private CKM_TextBox ckM_TextBox1;
    }
}