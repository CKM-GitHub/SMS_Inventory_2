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
    public partial class F0022 : Form
    {
        Login_Info LoginInfo;
        public F0022()
        {
            InitializeComponent();
        }

        public F0022(Login_Info LoginInfo)
        {
            InitializeComponent();
            this.LoginInfo = LoginInfo;
        }

        private void F0022_Load(object sender, EventArgs e)
        {

            this.reportViewer1.RefreshReport();
        }
    }
}
