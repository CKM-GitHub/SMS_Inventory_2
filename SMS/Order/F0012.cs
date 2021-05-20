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
    public partial class F0012 : SMS_Base
    {
        Login_Info loginInfo;

        public F0012()
        {
            InitializeComponent();
        }

        public F0012(Login_Info loginInfo)
            :base(loginInfo)
        {
            InitializeComponent();
            this.loginInfo = loginInfo;
            SetDesignerFunction();
        }

        private void SetDesignerFunction()
        {
            FunctionButtonDisabled(1);
            FunctionButtonDisabled(2);
            FunctionButtonDisabled(3);
            FunctionButtonDisabled(4);
            FunctionButtonDisabled(6);
            FunctionButtonDisabled(7);
            FunctionButtonDisabled(8);
            FunctionButtonDisabled(10);
            FunctionButtonDisabled(11);

            FormMode = "登録";
        }

        private void F0012_Load(object sender, EventArgs e)
        {
            FormName = "棚番入力";
            dgv1.DisabledColumn("colNo,Column4,Column5,Column6,Column7,Column8,Column9,Column10");
            chk1.Checked = true;
            BindCombo();
           // BindGrid();
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

        private void BindGrid()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("No");
            dt.Columns.Add("Storage");
            dt.Columns.Add("ExistingStorage");
            dt.Columns.Add("SKUCD");
            dt.Columns.Add("ProductName");
            dt.Columns.Add("Color");
            dt.Columns.Add("Size");
            dt.Columns.Add("JANCD");
            dt.Columns.Add("ReceivedNo");

            dt.Rows.Add("1", "XXXXXXXXXX", "XXXXXXXX", "XXXXXXXXX1XXXXXXXX2XXXXXXXX3", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸＸ40", "ＸＸＸＸＸＸＸＸＸ10", "ＸＸＸＸＸＸＸＸＸ10", "1234567890123", "99,999");
            dt.Rows.Add("2", "XXXXXXXXXX", "XXXXXXXX", "XXXXXXXXX1XXXXXXXX2XXXXXXXX3", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸＸ40", "ＸＸＸＸＸＸＸＸＸ10", "ＸＸＸＸＸＸＸＸＸ10", "1234567890123", "99,999");
            dt.Rows.Add("3", "XXXXXXXXXX", "XXXXXXXX", "XXXXXXXXX1XXXXXXXX2XXXXXXXX3", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸＸ40", "ＸＸＸＸＸＸＸＸＸ10", "ＸＸＸＸＸＸＸＸＸ10", "1234567890123", "99,999");
            dt.Rows.Add("4", "XXXXXXXXXX", "XXXXXXXX", "XXXXXXXXX1XXXXXXXX2XXXXXXXX3", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸＸ40", "ＸＸＸＸＸＸＸＸＸ10", "ＸＸＸＸＸＸＸＸＸ10", "1234567890123", "99,999");
            dt.Rows.Add("5", "XXXXXXXXXX", "XXXXXXXX", "XXXXXXXXX1XXXXXXXX2XXXXXXXX3", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸＸ40", "ＸＸＸＸＸＸＸＸＸ10", "ＸＸＸＸＸＸＸＸＸ10", "1234567890123", "99,999");
            dt.Rows.Add("6", "XXXXXXXXXX", "XXXXXXXX", "XXXXXXXXX1XXXXXXXX2XXXXXXXX3", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸＸ40", "ＸＸＸＸＸＸＸＸＸ10", "ＸＸＸＸＸＸＸＸＸ10", "1234567890123", "99,999");
            dt.Rows.Add("7", "XXXXXXXXXX", "XXXXXXXX", "XXXXXXXXX1XXXXXXXX2XXXXXXXX3", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸＸ40", "ＸＸＸＸＸＸＸＸＸ10", "ＸＸＸＸＸＸＸＸＸ10", "1234567890123", "99,999");
            dt.Rows.Add("8", "XXXXXXXXXX", "XXXXXXXX", "XXXXXXXXX1XXXXXXXX2XXXXXXXX3", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸＸ40", "ＸＸＸＸＸＸＸＸＸ10", "ＸＸＸＸＸＸＸＸＸ10", "1234567890123", "99,999");
            dt.Rows.Add("9", "XXXXXXXXXX", "XXXXXXXX", "XXXXXXXXX1XXXXXXXX2XXXXXXXX3", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸＸ40", "ＸＸＸＸＸＸＸＸＸ10", "ＸＸＸＸＸＸＸＸＸ10", "1234567890123", "99,999");
            dt.Rows.Add("10", "XXXXXXXXXX", "XXXXXXXX", "XXXXXXXXX1XXXXXXXX2XXXXXXXX3", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸＸ40", "ＸＸＸＸＸＸＸＸＸ10", "ＸＸＸＸＸＸＸＸＸ10", "1234567890123", "99,999");
            dt.Rows.Add("11", "XXXXXXXXXX", "XXXXXXXX", "XXXXXXXXX1XXXXXXXX2XXXXXXXX3", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸＸ40", "ＸＸＸＸＸＸＸＸＸ10", "ＸＸＸＸＸＸＸＸＸ10", "1234567890123", "99,999");
            dt.Rows.Add("12", "XXXXXXXXXX", "XXXXXXXX", "XXXXXXXXX1XXXXXXXX2XXXXXXXX3", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸＸ40", "ＸＸＸＸＸＸＸＸＸ10", "ＸＸＸＸＸＸＸＸＸ10", "1234567890123", "99,999");
            dt.Rows.Add("13", "XXXXXXXXXX", "XXXXXXXX", "XXXXXXXXX1XXXXXXXX2XXXXXXXX3", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸＸ40", "ＸＸＸＸＸＸＸＸＸ10", "ＸＸＸＸＸＸＸＸＸ10", "1234567890123", "99,999");
            dt.Rows.Add("14", "XXXXXXXXXX", "XXXXXXXX", "XXXXXXXXX1XXXXXXXX2XXXXXXXX3", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸＸ40", "ＸＸＸＸＸＸＸＸＸ10", "ＸＸＸＸＸＸＸＸＸ10", "1234567890123", "99,999");
            dt.Rows.Add("15", "XXXXXXXXXX", "XXXXXXXX", "XXXXXXXXX1XXXXXXXX2XXXXXXXX3", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸＸ40", "ＸＸＸＸＸＸＸＸＸ10", "ＸＸＸＸＸＸＸＸＸ10", "1234567890123", "99,999");
            dt.AcceptChanges();

            dgv1.AutoGenerateColumns = false;
            dgv1.DataSource = dt;

        }

        private void btnAllCheck_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dgv1.Rows.Count; i++)
            {
                dgv1[1, i].Value = true;//check every row in the forth column

            }
        }

        private void btnAllCancel_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dgv1.Rows.Count; i++)
            {
                dgv1[1, i].Value = false;//uncheck every row in the forth column

            }
        }
    }
}
