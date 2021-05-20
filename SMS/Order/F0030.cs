using SMS_Entity;
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
    public partial class F0030 : SMS_Base
    {
        Login_Info loginInfo;
        public F0030()
        {
            InitializeComponent();
        }


        public F0030(Login_Info loginInfo)
            : base(loginInfo)
        {

            InitializeComponent();
            SetDesignerFunction();
            this.loginInfo = loginInfo;

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
            checkBox1.Checked = true;
            checkBox3.Checked = true;
            checkBox5.Checked = true;
            checkBox13.Checked = true;
        }

        private void F0030_Load(object sender, EventArgs e)
        {
            SetDesignerFunction();
            FormName = "出荷指示書・ピッキングリスト印刷";
            BindCombo();
        }

        public override void functionClick(int btnIndex)
        {
            switch (btnIndex)
            {
                case 12:
                    F12();
                    break;
            }
        }

        private void F12()
        {
            frmF0018 frmf18 = new frmF0018();
            frmf18.Show();
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
            cboWarehouse.SelectedIndex = -1;

            cboWarehouse.DisplayMember = "warehouseName";
            cboWarehouse.ValueMember = "warehouseID";

            this.cboWarehouse.DataSource = dt;


        }
              
    }
}
