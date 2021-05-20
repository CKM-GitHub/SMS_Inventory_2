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
    public partial class temp : SMS_Base
    {
        public temp()
        {
            InitializeComponent();
            SetDesignerFunction();
        }

        private void SetDesignerFunction()
        {
            FunctionButtonDisabled(6);
            FunctionButtonDisabled(8);
        }

        private void F0005_Load(object sender, EventArgs e)
        {
            BindCombo();
            FormName = "WEB受注入力";
        }

        private void BindCombo()
        {
            
        }
    }
}
