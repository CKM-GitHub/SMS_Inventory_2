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
    public partial class F0029 : SMS_Base
    {
        Login_Info loginInfo;
        public F0029()
        {
            InitializeComponent();
        }

        public F0029(Login_Info loginInfo)
            : base(loginInfo)
        {

            InitializeComponent();
            SetDesignerFunction();
            this.loginInfo = loginInfo;

        }
       

        private void F0029_Load(object sender, EventArgs e)
        {
            FormName = "回答納期確認書印刷";
            BindCombo();
            BindDropdown();
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
        private void BindCombo()
        {
            #region 店舗

            DataTable dtstore = new DataTable();

            dtstore.Columns.Add("storeID");
            dtstore.Columns.Add("storeName");

            dtstore.Rows.Add("-1", "");
            dtstore.Rows.Add("1", "Web店");
            dtstore.Rows.Add("2", "豊中店");
            dtstore.Rows.Add("3", "石橋店");
            dtstore.Rows.Add("4", "三宮店");
            dtstore.Rows.Add("5", "江坂店");

            dtstore.AcceptChanges();
            comboBox2.SelectedIndex = -1;

            comboBox2.DisplayMember = "storeName";
            comboBox2.ValueMember = "storeID";

            comboBox2.DataSource = dtstore;
            #endregion
        }

        private void BindDropdown()
        {
            DataTable dt = new DataTable();

            dt.Columns.Add("ID");
            dt.Columns.Add("Name");

            dt.Rows.Add("-1", "");
            dt.Rows.Add("1", "確認中");

            dt.AcceptChanges();
            ddl1.SelectedIndex = -1;

            ddl1.DisplayMember = "Name";
            ddl1.ValueMember = "ID";

            ddl1.DataSource = dt;
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
            frmF0006 frmf6 = new frmF0006();
            frmf6.Show();
        }
    }
}
