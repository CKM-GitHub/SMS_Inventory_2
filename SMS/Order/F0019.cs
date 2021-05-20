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
using System.IO;

namespace SMS.Order
{
    public partial class F0019 : SMS_Base
    {
        Login_Info loginInfo;

        public F0019()
        {
            InitializeComponent();
        }

        public F0019(Login_Info loginInfo)
            :base(loginInfo)
        {
            InitializeComponent();
            this.loginInfo = loginInfo;
            SetDesignerFunction();
        }

        private void SetDesignerFunction()
        {
            lblMode.Visible = false;
            FunctionButtonDisabled(2);
            FunctionButtonDisabled(3);
            FunctionButtonDisabled(4);
            FunctionButtonDisabled(6);
            FunctionButtonDisabled(7);
            FunctionButtonDisabled(8);
            FunctionButtonDisabled(10);
            FunctionButtonDisabled(11);
        }

        private void F0019_Load(object sender, EventArgs e)
        {
            FormName = "送り状データ取込";
            BindCombo();
        }

        public void BindCombo()
        {
            DataTable dt = new DataTable();

            dt.Columns.Add("shippingID");
            dt.Columns.Add("shippingName");

            dt.Rows.Add("-1", "");
            dt.Rows.Add("1", "ヤマト運輸");
            dt.Rows.Add("2", "佐川急便");

            dt.AcceptChanges();
            cboShip.SelectedIndex = -1;

            cboShip.DisplayMember = "shippingName";
            cboShip.ValueMember = "shippingID";

            this.cboShip.DataSource = dt;
        }

        private void btnShowDialog_Click(object sender, EventArgs e)
        {
            string folderPath = "C:\\Csv\\";
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
                System.Diagnostics.Process.Start(folderPath);
            }
            else
            {
                System.Diagnostics.Process.Start(folderPath);
            }
        }
    }
}
