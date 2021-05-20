using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SMS.Order
{
    public partial class frmF0006 : Form
    {
        private SMS_Entity.Login_Info login_Info;

        public frmF0006()
        {
            InitializeComponent();
        }

        //public F0006(SMS_Entity.Login_Info login_Info)
        //{
        //    // TODO: Complete member initialization
        //    this.login_Info = login_Info;
        //}

        private void F0006_Load(object sender, EventArgs e)
        {
            
            this.reportViewer1.RefreshReport();
        }
    }
}
