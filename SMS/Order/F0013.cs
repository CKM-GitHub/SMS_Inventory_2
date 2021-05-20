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
    public partial class F0013 : SMS_Base
    {
        Login_Info loginInfo;

        public F0013()
        {
            InitializeComponent();
        }

        public F0013(Login_Info loginInfo)
            :base (loginInfo)
        {
            InitializeComponent();
            this.loginInfo=loginInfo;
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

        private void F0013_Load(object sender, EventArgs e)
        {
            FormName = "棚入れリスト印刷";
            BindCombo();
            chk1.Checked = chk3.Checked = true;
        }

        public void BindCombo()
        {
            DataTable dt = new DataTable();

            dt.Columns.Add("warehouseID");
            dt.Columns.Add("warehouseName");

            dt.Rows.Add("-1", "");
            dt.Rows.Add("1", "本社Web倉庫");
            dt.Rows.Add("2", "豊中店倉庫");
            dt.Rows.Add("3", "石橋店倉庫");
            dt.Rows.Add("4", "三宮店倉庫");
            dt.Rows.Add("5", "江坂店倉庫");

            dt.AcceptChanges();
            cbowarehouse.SelectedIndex = -1;

            cbowarehouse.DisplayMember = "warehouseName";
            cbowarehouse.ValueMember = "warehouseID";

            this.cbowarehouse.DataSource = dt;
        }
    }
}
