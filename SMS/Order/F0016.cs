using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data;
using SMS_Entity;

namespace SMS.Order
{
    public partial class F0016 : SMS_Base
    {
        Login_Info loginInfo;

        public F0016()
        {
            InitializeComponent();
        }
        public F0016(Login_Info loginInfo) 
            : base (loginInfo)
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
            FunctionButtonDisabled(12);
        }

        private void F0016_Load(object sender, EventArgs e)
        {
            FormName = "仕入照会(仕入明細)";
            dgv.DisabledColumn("*");

            BindCombo();
            //BindGrid();
        }

        private void BindCombo()
        {
            DataTable dtStore = new DataTable();

            dtStore.Columns.Add("storeID");
            dtStore.Columns.Add("storeName");

            dtStore.Rows.Add("-1", "");
            dtStore.Rows.Add("0", "Web店");
            dtStore.Rows.Add("1", "豊中店");
            dtStore.Rows.Add("2", "石橋店");
            dtStore.Rows.Add("3", "三宮店");
            dtStore.Rows.Add("4", "江坂店");
            dtStore.Rows.Add("5", "楽天ラケットプラザ");
            dtStore.Rows.Add("6", "Yahooラケットプラザ");

            dtStore.AcceptChanges();
            cboStore.SelectedIndex = -1;

            this.cboStore.DisplayMember = "StoreName";
            this.cboStore.ValueMember = "StoreID";
            this.cboStore.DataSource = dtStore;

        }

        //private void BindGrid()
        //{
        //    DataTable dtgrid = new DataTable();

        //    dtgrid.Columns.Add("No.");
        //    dtgrid.Columns.Add("Button");
        //    dtgrid.Columns.Add("PurchaseNum");
        //    dtgrid.Columns.Add("PurchaseDate");
        //    dtgrid.Columns.Add("Suppliers");
        //    dtgrid.Columns.Add("SKUCD");
        //    dtgrid.Columns.Add("JANCD");
        //    dtgrid.Columns.Add("ItemCD");
        //    dtgrid.Columns.Add("ProductName");
        //    dtgrid.Columns.Add("ColorSize");
        //    dtgrid.Columns.Add("Remarks");
        //    dtgrid.Columns.Add("NumOfPurchase");
        //    dtgrid.Columns.Add("PurchasePrice");
        //    dtgrid.Columns.Add("PurchaseAmount");
        //    dtgrid.Columns.Add("NumOfOrder");
        //    dtgrid.Columns.Add("OrderUnitPrice");
        //    dtgrid.Columns.Add("OrderAmount");
        //    dtgrid.Columns.Add("OrderDate");
        //    dtgrid.Columns.Add("OrderNum");
        //    dtgrid.Columns.Add("ArrivalDay");
        //    dtgrid.Columns.Add("Supplier");
        //    dtgrid.Columns.Add("Store");
        //    dtgrid.Columns.Add("StaffInCharge");
        //    dtgrid.Columns.Add("PaymentDueDate");
        //    dtgrid.Columns.Add("PaymentDate");
        //    dtgrid.Columns.Add("InvoiceNum");

        //    string[] row1 = new string[] { "1", "", "XXXXXXXXX11", "2018/4/21", "XXXXXXXXXX ＸＸＸＸＸ10ＸＸＸＸＸＸ20", "XXXXXXXXX1XXXXXXXX3", "1234567890123", "XXXXXXXXX1XXXXXXXXX3", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸＸ40", "カラーサイズ等ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸＸ40", "社外備考・コメントＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸＸ40", "99,999", "9,999,999", "9,999,999", "99,999", "9,999,999", "9,999,999", "2018/4/21", "XXXXXXXXX11", "2018/4/21", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "ＸＸＸＸＸＸＸＸＸＸ	", "ＸＸＸＸＸＸＸＸＸＸ	", "2018/4/21", "2018/4/21", "XXXXXXXXX15" };
        //    string[] row2 = new string[] { "2", "", "XXXXXXXXX11", "2018/4/21", "XXXXXXXXXX ＸＸＸＸＸ10ＸＸＸＸＸＸ20", "XXXXXXXXX1XXXXXXXX3", "1234567890123", "XXXXXXXXX1XXXXXXXXX3", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸＸ40", "カラーサイズ等ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸＸ40", "社外備考・コメントＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸＸ40", "99,999", "9,999,999", "9,999,999", "99,999", "9,999,999", "9,999,999", "2018/4/21", "XXXXXXXXX11", "2018/4/21", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "ＸＸＸＸＸＸＸＸＸＸ	", "ＸＸＸＸＸＸＸＸＸＸ	", "2018/4/21", "2018/4/21", "XXXXXXXXX15" };
        //    string[] row3 = new string[] { "3", "", "XXXXXXXXX11", "2018/4/21", "XXXXXXXXXX ＸＸＸＸＸ10ＸＸＸＸＸＸ20", "XXXXXXXXX1XXXXXXXX3", "1234567890123", "XXXXXXXXX1XXXXXXXXX3", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸＸ40", "カラーサイズ等ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸＸ40", "社外備考・コメントＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸＸ40", "99,999", "9,999,999", "9,999,999", "99,999", "9,999,999", "9,999,999", "2018/4/21", "XXXXXXXXX11", "2018/4/21", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "ＸＸＸＸＸＸＸＸＸＸ	", "ＸＸＸＸＸＸＸＸＸＸ	", "2018/4/21", "2018/4/21", "XXXXXXXXX15" };
        //    string[] row4 = new string[] { "4", "", "XXXXXXXXX11", "2018/4/21", "XXXXXXXXXX ＸＸＸＸＸ10ＸＸＸＸＸＸ20", "XXXXXXXXX1XXXXXXXX3", "1234567890123", "XXXXXXXXX1XXXXXXXXX3", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸＸ40", "カラーサイズ等ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸＸ40", "社外備考・コメントＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸＸ40", "99,999", "9,999,999", "9,999,999", "99,999", "9,999,999", "9,999,999", "2018/4/21", "XXXXXXXXX11", "2018/4/21", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "ＸＸＸＸＸＸＸＸＸＸ	", "ＸＸＸＸＸＸＸＸＸＸ	", "2018/4/21", "2018/4/21", "XXXXXXXXX15" };
        //    string[] row5 = new string[] { "5", "", "XXXXXXXXX11", "2018/4/21", "XXXXXXXXXX ＸＸＸＸＸ10ＸＸＸＸＸＸ20", "XXXXXXXXX1XXXXXXXX3", "1234567890123", "XXXXXXXXX1XXXXXXXXX3", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸＸ40", "カラーサイズ等ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸＸ40", "社外備考・コメントＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸＸ40", "99,999", "9,999,999", "9,999,999", "99,999", "9,999,999", "9,999,999", "2018/4/21", "XXXXXXXXX11", "2018/4/21", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "ＸＸＸＸＸＸＸＸＸＸ	", "ＸＸＸＸＸＸＸＸＸＸ	", "2018/4/21", "2018/4/21", "XXXXXXXXX15" };
          

        //    dtgrid.Rows.Add(row1);
        //    dtgrid.Rows.Add(row2);
        //    dtgrid.Rows.Add(row3);
        //    dtgrid.Rows.Add(row4);
        //    dtgrid.Rows.Add(row5);

        //    dgv.AutoGenerateColumns = false;
        //    dgv.DataSource = dtgrid;



        //}

        private void smS_TextBox2_TextChanged(object sender, EventArgs e)
        {

        }

    }
}
