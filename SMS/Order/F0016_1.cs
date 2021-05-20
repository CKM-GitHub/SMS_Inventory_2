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
    public partial class F0016_1 : SMS_Base
    {
        public F0016_1()
        {
            InitializeComponent();
            SetDesignerFunction();
        }

        private void SetDesignerFunction()
        {
            FunctionButtonDisabled(6);
            FunctionButtonDisabled(7);
            FunctionButtonDisabled(8);
            FunctionButtonDisabled(10);
            FunctionButtonDisabled(11);
        }

        private void F0016_Load(object sender, EventArgs e)
        {
            FormName = "仕入入力（仕入から）";
            BindCombo();
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
        }
    }
}
