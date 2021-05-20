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
    public partial class F0014 : SMS_Base
    {
        public F0014()
        {
            InitializeComponent();
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
            FunctionButtonDisabled(9);
            FunctionButtonDisabled(10);
            FunctionButtonDisabled(11);
        }

        private void F0014_Load(object sender, EventArgs e)
        {
            FormName = "仕入単価訂正依頼書印刷";
            SetDesignerFunction();
            BindCombo();
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
        }
    }
}
