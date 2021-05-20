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
    public partial class F0010 : SMS_Base
    {
        Login_Info loginInfo; 

        public F0010()
        {
            InitializeComponent();           
        }

        public F0010(Login_Info loginInfo)
            : base(loginInfo)
        {
            InitializeComponent();
            this.loginInfo = loginInfo;
            SetDesignerFunction();
        }

        private void SetDesignerFunction()
        {
            lblMode.Visible = false;

            dgv1.DisabledColumn("*");

            FunctionButtonDisabled(2);
            FunctionButtonDisabled(3);
            FunctionButtonDisabled(4);
            FunctionButtonDisabled(6);
            FunctionButtonDisabled(7);
            FunctionButtonDisabled(8);
            FunctionButtonDisabled(12);
        }

        private void F0010_Load(object sender, EventArgs e)
        {
            FormName = "発注照会";
            BindCombo();
            //BindGrid();
        }

        private void BindCombo()
        {
            #region 店舗
            DataTable dtStore = new DataTable();
            dtStore.Columns.Add("StoreID");
            dtStore.Columns.Add("StoreName");

            dtStore.Rows.Add("-1", "");
            dtStore.Rows.Add("0", "Web店");
            dtStore.Rows.Add("1", "豊中店");
            dtStore.Rows.Add("2", "石橋店");
            dtStore.Rows.Add("3", "三宮店");
            dtStore.Rows.Add("4", "江坂店");

            dtStore.AcceptChanges();
            cboStore.SelectedIndex = -1;

            this.cboStore.DisplayMember = "StoreName";
            this.cboStore.ValueMember = "StoreID";
            this.cboStore.DataSource = dtStore;
            #endregion

            #region  倉庫
            DataTable dt = new DataTable();
            dt.Columns.Add("ID");
            dt.Columns.Add("Name");

            dt.Rows.Add("-1", "");
            dt.Rows.Add("0", "本社Web倉庫");
            dt.Rows.Add("1", "豊中店倉庫");
            dt.Rows.Add("2", "石橋店倉庫");
            dt.Rows.Add("3", "三宮店倉庫");
            dt.Rows.Add("4", "江坂店倉庫");

            dt.AcceptChanges();
            cboWarehouse.SelectedIndex = -1;

            this.cboWarehouse.DisplayMember = "Name";
            this.cboWarehouse.ValueMember = "ID";
            this.cboWarehouse.DataSource = dt;
            #endregion
        }

        private void BindGrid()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("No");
            dt.Columns.Add("OrderNo1");
            dt.Columns.Add("OrderDate");
            dt.Columns.Add("Supplier");
            dt.Columns.Add("SKUCD");
            dt.Columns.Add("JANCD");
            dt.Columns.Add("ItemCD");
            dt.Columns.Add("ProductName");
            dt.Columns.Add("ColorSize");
            dt.Columns.Add("Remark");

            dt.Columns.Add("StockDate");
            dt.Columns.Add("ArrivalDay");
            dt.Columns.Add("PurchaseDate");
            dt.Columns.Add("OrderNo2");
            dt.Columns.Add("Nōnyū-saki");
            dt.Columns.Add("Qty");
            dt.Columns.Add("OrderUnitPrice");
            dt.Columns.Add("OrderAmount");
            dt.Columns.Add("OrderMethod");
            dt.Columns.Add("Delivery");

            dt.Columns.Add("Status");
            dt.Columns.Add("Store");
            dt.Columns.Add("Staff");
            dt.Columns.Add("OrderNo3");


            dt.Rows.Add("1", "XXXXXXXXX11", "2018/4/21", "XXXXXXXXXX ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "XXXXXXXXX1XXXXXXXX2XXXXXXXXX3", "1234567890123", "XXXXXXXXX1XXXXXXXX2XXXXXXXXX3", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸＸ40", "カラーサイズ等ＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "社外備考・コメント10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "2018/4/21", "2018/4/21", "2018/4/21", "XXXXXXXXX11", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "99,999", "9,999,999", "9,999,999", "NET", "", "承認済", "ＸＸＸＸＸＸＸＸＸＸ", "ＸＸＸＸＸＸＸＸＸＸ", "XXXXXXXXX11");
            dt.Rows.Add("2", "XXXXXXXXX11", "2018/4/21", "XXXXXXXXXX ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "XXXXXXXXX1XXXXXXXX2XXXXXXXXX3", "1234567890123", "XXXXXXXXX1XXXXXXXX2XXXXXXXXX3", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸＸ40", "カラーサイズ等ＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "社外備考・コメント10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "2018/4/21", "2018/4/21", "2018/4/21", "XXXXXXXXX11", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "99,999", "9,999,999", "9,999,999", "FAX", "〇", "承認済", "ＸＸＸＸＸＸＸＸＸＸ", "ＸＸＸＸＸＸＸＸＸＸ", "XXXXXXXXX11");
            dt.Rows.Add("3", "XXXXXXXXX11", "2018/4/21", "XXXXXXXXXX ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "XXXXXXXXX1XXXXXXXX2XXXXXXXXX3", "1234567890123", "XXXXXXXXX1XXXXXXXX2XXXXXXXXX3", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸＸ40", "カラーサイズ等ＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "社外備考・コメント10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "2018/4/21", "2018/4/21", "2018/4/21", "XXXXXXXXX11", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "99,999", "9,999,999", "9,999,999", "NET", "", "承認済", "ＸＸＸＸＸＸＸＸＸＸ", "ＸＸＸＸＸＸＸＸＸＸ", "XXXXXXXXX11");
            dt.Rows.Add("4", "XXXXXXXXX11", "2018/4/21", "XXXXXXXXXX ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "XXXXXXXXX1XXXXXXXX2XXXXXXXXX3", "1234567890123", "XXXXXXXXX1XXXXXXXX2XXXXXXXXX3", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸＸ40", "カラーサイズ等ＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "社外備考・コメント10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "2018/4/21", "2018/4/21", "2018/4/21", "XXXXXXXXX11", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "99,999", "9,999,999", "9,999,999", "EDI", "〇", "承認済", "ＸＸＸＸＸＸＸＸＸＸ", "ＸＸＸＸＸＸＸＸＸＸ", "XXXXXXXXX11");
            dt.Rows.Add("5", "XXXXXXXXX11", "2018/4/21", "XXXXXXXXXX ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "XXXXXXXXX1XXXXXXXX2XXXXXXXXX3", "1234567890123", "XXXXXXXXX1XXXXXXXX2XXXXXXXXX3", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸＸ40", "カラーサイズ等ＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "社外備考・コメント10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "2018/4/21", "2018/4/21", "2018/4/21", "XXXXXXXXX11", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "99,999", "9,999,999", "9,999,999", "EDI", "", "承認済", "ＸＸＸＸＸＸＸＸＸＸ", "ＸＸＸＸＸＸＸＸＸＸ", "XXXXXXXXX11");
           
            dt.AcceptChanges();

            dgv1.AutoGenerateColumns = false;
            dgv1.DataSource = dt;

        }
    }
}
