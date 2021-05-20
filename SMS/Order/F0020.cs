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
    public partial class F0020 : SMS_Base
    {
        Login_Info loginInfo;

        public F0020()
        {
            InitializeComponent();
        }

        public F0020(Login_Info loginInfo)
            : base(loginInfo)
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
            FormMode = "登録";
            dgvF0020.DisabledColumn("colNo,colCostPerUnit,colGrossProfit,colSet,colSaleTax,colCostPrice,colTax,colPersentage");
        }

        private void F0020_Load(object sender, EventArgs e)
        {
            FormName = "見積入力";
            BindCombo();
            //BindGrid();
        }

        private void BindCombo()
        {
            #region conbo store
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

            #region conbo order
            DataTable dtOrder = new DataTable();
            dtOrder.Columns.Add("OrderID");
            dtOrder.Columns.Add("OrderName");

            dtOrder.Rows.Add("-1", "");
            dtOrder.Rows.Add("0", "Ａ：確実");
            dtOrder.Rows.Add("1", "Ｂ：ほぼ");
            dtOrder.Rows.Add("2", "Ｃ：厳しい");
            dtOrder.Rows.Add("3", "Ｄ：無理");

            dtOrder.AcceptChanges();
            cboOrder.SelectedIndex = -1;

            this.cboOrder.DisplayMember = "OrderName";
            this.cboOrder.ValueMember = "OrderID";
            this.cboOrder.DataSource = dtOrder;
            #endregion

        }

        //private void BindGrid()
        //{
        //    DataTable dt = new DataTable();
        //    dt.Columns.Add("No");
        //    dt.Columns.Add("ProductName");
        //    dt.Columns.Add("Color");
        //    dt.Columns.Add("Size");
        //    dt.Columns.Add("Set");
        //    dt.Columns.Add("EstimatedNo");
        //    dt.Columns.Add("Unit");
        //    dt.Columns.Add("UnitSellingPrice");
        //    dt.Columns.Add("SaleAmountTax");
        //    dt.Columns.Add("SaleTax");

        //    dt.Columns.Add("CostPerUnit");
        //    dt.Columns.Add("CostPrice");
        //    dt.Columns.Add("InternalRemark");
        //    dt.Columns.Add("ExternalRemark");
        //    dt.Columns.Add("IndividualDistributor");
        //    dt.Columns.Add("GrossProfit");
        //    dt.Columns.Add("Tax");
        //    dt.Columns.Add("Persentage");
        //    //dt.Columns.Add("UnitPrice");
        //    //dt.Columns.Add("ReductionAmount");
        //    //dt.Columns.Add("PurchaseAmount");

        //    dt.Rows.Add("1", "1234567890123･･･ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ4", "ＸＸＸＸＸＸＸＸＸ10", "ＸＸＸＸＸＸＸＸＸ10", "〇", "-99,999", "ＸＸＸ", "99,999,999", "-999,999,999", "-999,999,999", "99,999,999", "-999,999,999", "社内備考・コメント10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "社外備考・コメント10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "個別販売先等ＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "-999,999,999", "税込", "8%");
        //    dt.Rows.Add("2", "1234567890123･･･ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ4", "ＸＸＸＸＸＸＸＸＸ10", "ＸＸＸＸＸＸＸＸＸ10", "〇", "-99,999", "ＸＸＸ", "99,999,999", "-999,999,999", "-999,999,999", "99,999,999", "-999,999,999", "社内備考・コメント10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "社外備考・コメント10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "個別販売先等ＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "-999,999,999", "税込", "8%");
        //    dt.Rows.Add("3", "1234567890123･･･ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ4", "ＸＸＸＸＸＸＸＸＸ10", "ＸＸＸＸＸＸＸＸＸ10", "〇", "-99,999", "ＸＸＸ", "99,999,999", "-999,999,999", "-999,999,999", "99,999,999", "-999,999,999", "社内備考・コメント10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "社外備考・コメント10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "個別販売先等ＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "-999,999,999", "税込", "8%");
        //    dt.Rows.Add("4", "1234567890123･･･ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ4", "ＸＸＸＸＸＸＸＸＸ10", "ＸＸＸＸＸＸＸＸＸ10", "〇", "-99,999", "ＸＸＸ", "99,999,999", "-999,999,999", "-999,999,999", "99,999,999", "-999,999,999", "社内備考・コメント10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "社外備考・コメント10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "個別販売先等ＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "-999,999,999", "税込", "8%");
        //    dt.Rows.Add("5", "1234567890123･･･ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ4", "ＸＸＸＸＸＸＸＸＸ10", "ＸＸＸＸＸＸＸＸＸ10", "〇", "-99,999", "ＸＸＸ", "99,999,999", "-999,999,999", "-999,999,999", "99,999,999", "-999,999,999", "社内備考・コメント10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "社外備考・コメント10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "個別販売先等ＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "-999,999,999", "税込", "8%");
        //    dt.Rows.Add("6", "1234567890123･･･ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ4", "ＸＸＸＸＸＸＸＸＸ10", "ＸＸＸＸＸＸＸＸＸ10", "〇", "-99,999", "ＸＸＸ", "99,999,999", "-999,999,999", "-999,999,999", "99,999,999", "-999,999,999", "社内備考・コメント10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "社外備考・コメント10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "個別販売先等ＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "-999,999,999", "税込", "8%");

        //    dt.AcceptChanges();
        //    dgvF0020.AutoGenerateColumns = false;
        //    dgvF0020.DataSource = dt;

        //    smS_TextBox11.Text = "見積書下部の備考欄に印刷するＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸ35ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸ35ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸ35ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸ35ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸ35ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸ35";
        //}

        public override void functionClick(int btnIndex)
        {
            switch (btnIndex)
            {
                case 11:
                    F11();
                    break;
            }
        }

        private void F11()
        {
            frmF0021 frmf21 = new frmF0021();
            frmf21.Show();
        }
    }
}
