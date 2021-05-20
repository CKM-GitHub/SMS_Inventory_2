using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SMS_Entity;

namespace SMS.Order
{
    public partial class frmF0021 : Form
    {
        Login_Info LoginInfo;
        public frmF0021()
        {
            InitializeComponent();
        }

        public frmF0021(Login_Info LoginInfo)
        {
            InitializeComponent();
            this.LoginInfo = LoginInfo;
        }

        private void F0021_Load(object sender, EventArgs e)
        {

            this.reportViewer1.RefreshReport();
        }
    }
}
