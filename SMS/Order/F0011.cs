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
    public partial class F0011 : SMS_Base
    {
        Login_Info loginInfo;

        public F0011()
        {
            InitializeComponent();
        }

        public F0011(Login_Info loginInfo)
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

        private void F0011_Load(object sender, EventArgs e)
        {
            FormName = "入荷照会";
            //BindGrid();
            BindCombo();
        }


        private void BindGrid()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("No");
            dt.Columns.Add("ArrivalDay");
            dt.Columns.Add("StockDate");
            dt.Columns.Add("PurchaseDate");
            dt.Columns.Add("IncomeGoods");
            dt.Columns.Add("SKUCD");
            dt.Columns.Add("JANCD");
            dt.Columns.Add("ProductName");
            dt.Columns.Add("Color");
            dt.Columns.Add("Size");

            dt.Columns.Add("ExceptedNo");
            dt.Columns.Add("ReceivedNo");
            dt.Columns.Add("Supplier");
            dt.Columns.Add("Warehouse");
            dt.Columns.Add("Delivery");
            dt.Columns.Add("OrderNo1");
            dt.Columns.Add("OrderNo2");
            dt.Columns.Add("InboundNo");
            dt.Columns.Add("PurchaseNo");
            dt.Columns.Add("InvoiceNo");

            dt.Rows.Add("1", "2018/4/21", "2018/4/21", "2018/4/21", "発注", "XXXXXXXXX1XXXXXXXX2XXXXXXXX3", "1234567890123", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸＸ40", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30", "99,999", "99,999", "XXXXXXXXXX ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "〇", "XXXXXXXXX11", "XXXXXXXXX11", "XXXXXXXXX11", "XXXXXXXXX11", "XXXXXXXXX15");
            dt.Rows.Add("2", "", "2018/4/21", "", "発注", "XXXXXXXXX1XXXXXXXX2XXXXXXXX3", "1234567890123", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸＸ40", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30", "99,999", "99,999", "XXXXXXXXXX ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "", "XXXXXXXXX11", "XXXXXXXXX11", "XXXXXXXXX11", "XXXXXXXXX11", "XXXXXXXXX15");
            dt.Rows.Add("3", "", "2018/4/21", "", "発注", "XXXXXXXXX1XXXXXXXX2XXXXXXXX3", "1234567890123", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸＸ40", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30", "99,999", "99,999", "XXXXXXXXXX ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "〇", "XXXXXXXXX11", "XXXXXXXXX11", "XXXXXXXXX11", "XXXXXXXXX11", "XXXXXXXXX15");
            dt.Rows.Add("4", "2018/4/21", "2018/4/21", "2018/4/21", "発注", "XXXXXXXXX1XXXXXXXX2XXXXXXXX3", "1234567890123", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸＸ40", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30", "99,999", "99,999", "XXXXXXXXXX ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "", "XXXXXXXXX11", "XXXXXXXXX11", "XXXXXXXXX11", "XXXXXXXXX11", "XXXXXXXXX15");
            dt.Rows.Add("5", "2018/4/21", "2018/4/21", "2018/4/21", "発注", "XXXXXXXXX1XXXXXXXX2XXXXXXXX3", "1234567890123", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸＸ40", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸＸ30", "99,999", "99,999", "XXXXXXXXXX ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "", "XXXXXXXXX11", "XXXXXXXXX11", "XXXXXXXXX11", "XXXXXXXXX11", "XXXXXXXXX15");

            dt.AcceptChanges();

            dgv1.AutoGenerateColumns = false;
            dgv1.DataSource = dt;

        }

        private void BindCombo()
        {
            #region 倉庫

            DataTable dtwarehouse = new DataTable();

            dtwarehouse.Columns.Add("warehouseID");
            dtwarehouse.Columns.Add("warehouseName");

            dtwarehouse.Rows.Add("-1", "");
            dtwarehouse.Rows.Add("1", "本社Web倉庫");
            dtwarehouse.Rows.Add("2", "豊中店倉庫");
            dtwarehouse.Rows.Add("3", "石橋店倉庫");
            dtwarehouse.Rows.Add("4", "三宮店倉庫");
            dtwarehouse.Rows.Add("5", "江坂店倉庫");

            dtwarehouse.AcceptChanges();
            comboBox1.SelectedIndex = -1;

            comboBox1.DisplayMember = "warehouseName";
            comboBox1.ValueMember = "warehouseID";

            comboBox1.DataSource = dtwarehouse;
            #endregion

            #region 移動元倉庫

            DataTable dtsource = new DataTable();

            dtsource.Columns.Add("sourceID");
            dtsource.Columns.Add("sourceName");

            dtsource.Rows.Add("-1", "");
            dtsource.Rows.Add("1", "本社Web倉庫");
            dtsource.Rows.Add("2", "豊中店倉庫");
            dtsource.Rows.Add("3", "石橋店倉庫");
            dtsource.Rows.Add("4", "三宮店倉庫");
            dtsource.Rows.Add("5", "江坂店倉庫");

            dtsource.AcceptChanges();
            comboBox2.SelectedIndex = -1;

            comboBox2.DisplayMember = "sourceName";
            comboBox2.ValueMember = "sourceID";

            comboBox2.DataSource = dtsource;

            #endregion


        }
    }
}
