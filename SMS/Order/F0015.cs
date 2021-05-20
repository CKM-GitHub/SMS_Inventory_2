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
    public partial class F0015 : SMS_Base
    {
        Login_Info loginInfo;

        public F0015()
        {
            InitializeComponent();
        }

        public F0015(Login_Info loginInfo)
            :base(loginInfo)
        {
            InitializeComponent();
            this.loginInfo = loginInfo;
            SetDesignerFunction();
        }

        private void SetDesignerFunction()
        {
            FunctionButtonDisabled(6);
            FunctionButtonDisabled(7);
            FunctionButtonDisabled(8);
            FunctionButtonDisabled(10);
            FunctionButtonDisabled(11);
            
            dgvF0015.EnabledColumn("colchk,colUnitPrice,colReductionAmount");

            FormMode = "登録";
        }

        private void F0015_Load(object sender, EventArgs e)
        {
            FormName = "仕入入力（入荷から）";
            BindCombo();
            //BindGrid();
        }

        public void BindCombo()
        {
            #region 店舗
            DataTable dt = new DataTable();

            dt.Columns.Add("storeID");
            dt.Columns.Add("storeName");

            dt.Rows.Add("-1", "");
            dt.Rows.Add("1", "Web店");
            dt.Rows.Add("2", "豊中店");
            dt.Rows.Add("3", "石橋店");
            dt.Rows.Add("4", "三宮店");
            dt.Rows.Add("5", "江坂店");
            dt.Rows.Add("6", "楽天ラケットプラザ");
            dt.Rows.Add("7", "Yahooラケットプラザ");

            dt.AcceptChanges();
            cboStore.SelectedIndex = -1;

            cboStore.DisplayMember = "storeName";
            cboStore.ValueMember = "storeID";

            this.cboStore.DataSource = dt;
            #endregion

            #region 自社倉庫

            DataTable dt2 = new DataTable();

            dt2.Columns.Add("warehouseID");
            dt2.Columns.Add("warehouseName");

            dt2.Rows.Add("-1", "");
            dt2.Rows.Add("1", "本社Web倉庫");
            dt2.Rows.Add("1", "豊中店倉庫");
            dt2.Rows.Add("1", "石橋店倉庫");
            dt2.Rows.Add("1", "三宮店倉庫");
            dt2.Rows.Add("1", "江坂店倉庫");

            dt2.AcceptChanges();
            cboWarehouse.SelectedIndex = -1;

            cboWarehouse.DisplayMember = "warehouseName";
            cboWarehouse.ValueMember = "warehouseID";

            this.cboWarehouse.DataSource = dt2;
            #endregion
        }

        private void BindGrid()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("No");
            dt.Columns.Add("PurchaseNo");
            dt.Columns.Add("Delivery");
            dt.Columns.Add("ArrivalDay");
            dt.Columns.Add("InboundNo");
            dt.Columns.Add("OrderNo1");
            dt.Columns.Add("Supplier");
            dt.Columns.Add("SKUCD");
            dt.Columns.Add("JANCD");
            dt.Columns.Add("ProductName");

            dt.Columns.Add("InvoiceNo");
            dt.Columns.Add("InternalRemark");
            dt.Columns.Add("ExternalRemark");
            dt.Columns.Add("OrderNo2");
            dt.Columns.Add("OrderUnitPrice");
            dt.Columns.Add("PurchaseOrder");
            dt.Columns.Add("PurchaseCount");
            dt.Columns.Add("Unit");
            dt.Columns.Add("UnitPrice");
            dt.Columns.Add("ReductionAmount");

            dt.Columns.Add("PurchaseAmount");
            dt.Columns.Add("Tax");

            dt.Rows.Add("1", "XXXXXXXXX11", "", "9999/99/99", "XXXXXXXXX11", "XXXXXXXXX11", "ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ", "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3", "XXXXXXXXX1XX3", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "XXXXXXXX10XXXX5", "社内備考・コメント10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "社外備考・コメント10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "-99,999", "-9,999,999", "-9,999,999", "-99,999", "ＸＸＸ", "-9,999,999", "-9,999,999", "-9,999,999", "課税");
            dt.Rows.Add("2", "XXXXXXXXX11", "〇", "9999/99/99", "XXXXXXXXX11", "XXXXXXXXX11", "ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ", "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3", "XXXXXXXXX1XX3", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "XXXXXXXX10XXXX5", "社内備考・コメント10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "社外備考・コメント10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "-99,999", "-9,999,999", "-9,999,999", "-99,999", "ＸＸＸ", "-9,999,999", "-9,999,999", "-9,999,999", "課税");
            dt.Rows.Add("3", "XXXXXXXXX11", "", "9999/99/99", "XXXXXXXXX11", "XXXXXXXXX11", "ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ", "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3", "XXXXXXXXX1XX3", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "XXXXXXXX10XXXX5", "社内備考・コメント10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "社外備考・コメント10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "-99,999", "-9,999,999", "-9,999,999", "-99,999", "ＸＸＸ", "-9,999,999", "-9,999,999", "-9,999,999", "課税");
            dt.Rows.Add("4", "XXXXXXXXX11", "", "9999/99/99", "XXXXXXXXX11", "XXXXXXXXX11", "ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ", "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3", "XXXXXXXXX1XX3", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "XXXXXXXX10XXXX5", "社内備考・コメント10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "社外備考・コメント10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "-99,999", "-9,999,999", "-9,999,999", "-99,999", "ＸＸＸ", "-9,999,999", "-9,999,999", "-9,999,999", "課税");
            dt.Rows.Add("5", "XXXXXXXXX11", "", "9999/99/99", "XXXXXXXXX11", "XXXXXXXXX11", "ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ", "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3", "XXXXXXXXX1XX3", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "XXXXXXXX10XXXX5", "社内備考・コメント10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "社外備考・コメント10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "-99,999", "-9,999,999", "-9,999,999", "-99,999", "ＸＸＸ", "-9,999,999", "-9,999,999", "-9,999,999", "課税");
            dt.Rows.Add("6", "XXXXXXXXX11", "〇", "9999/99/99", "XXXXXXXXX11", "XXXXXXXXX11", "ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ", "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3", "XXXXXXXXX1XX3", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "XXXXXXXX10XXXX5", "社内備考・コメント10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "社外備考・コメント10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "-99,999", "-9,999,999", "-9,999,999", "-99,999", "ＸＸＸ", "-9,999,999", "-9,999,999", "-9,999,999", "課税");

            dt.AcceptChanges();
            dgvF0015.AutoGenerateColumns = false;
            dgvF0015.DataSource = dt;
        }

        private void btnCheckAll_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dgvF0015.Rows.Count; i++)
            {

                dgvF0015[1, i].Value = true;//check every row in the forth column

            }
        }

        private void btnUnCheckAll_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dgvF0015.Rows.Count; i++)
            {

                dgvF0015[1, i].Value = false;//check every row in the forth column

            }

        }
    }
}
