namespace CKM_Controls
{
    partial class CKM_Label
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
            this.lblCKM = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // lblCKM
            // 
            this.lblCKM.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(246)))), ((int)(((byte)(228)))), ((int)(((byte)(204)))));
            this.lblCKM.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblCKM.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblCKM.Location = new System.Drawing.Point(0, 0);
            this.lblCKM.Name = "lblCKM";
            this.lblCKM.Size = new System.Drawing.Size(69, 18);
            this.lblCKM.TabIndex = 0;
            this.lblCKM.Text = "label1";
            this.lblCKM.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // CKM_Label
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.lblCKM);
            this.Font = new System.Drawing.Font("MS Gothic", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "CKM_Label";
            this.Size = new System.Drawing.Size(69, 18);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label lblCKM;
    }
}
