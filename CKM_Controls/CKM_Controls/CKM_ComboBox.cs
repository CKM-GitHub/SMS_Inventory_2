using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Collections;
using System.Drawing;

namespace CKM_Controls
{
    public partial class CKM_ComboBox : ComboBox
    {
    
        //public CKM_ComboBox()
        //{
        //    this.Enter+=CKM_ComboBox_Enter;
        //    InitializeComponent();
        //}

        public CKM_ComboBox(IContainer container)
        {
            container.Add(this);

            InitializeComponent();
            this.BackColor = Color.White;
            this.Enter += CKM_ComboBox_Enter;
            this.Leave+=CKM_ComboBox_Leave;
            //this.ComboBox1.DropDownStyle = ComboBoxStyle.DropDownList;
            this.DropDownStyle = ComboBoxStyle.DropDownList;
           
        }

        private void CKM_ComboBox_Enter(object sender, EventArgs e)
        {
            this.BackColor = Color.LightYellow;
            
        }

       
        private void CKM_ComboBox_Leave(object sender, EventArgs e)
        {
            this.BackColor = Color.White;
        }
        
    }
}
