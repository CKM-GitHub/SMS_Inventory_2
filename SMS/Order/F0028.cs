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
    public partial class frmF0028 : SMS_Base
    {
        Login_Info loginInfo;

        public frmF0028()
        {
            InitializeComponent();
        }

        public frmF0028(Login_Info loginInfo)
            :base(loginInfo)
        {
            InitializeComponent();
            this.loginInfo = loginInfo;
            SetDesignerFunction();
        }

        private void SetDesignerFunction()
        {
            lblMode.Visible = false;
            FormName = "発注書印刷";
            FunctionButtonDisabled(2);
            FunctionButtonDisabled(3);
            FunctionButtonDisabled(4);
            FunctionButtonDisabled(6);
            FunctionButtonDisabled(7);
            FunctionButtonDisabled(8);
            FunctionButtonDisabled(10);
            FunctionButtonDisabled(11);
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
            frmF0009 frmf9 = new frmF0009();
            frmf9.Show();
        }

        private void frmF0028_Load(object sender, EventArgs e)
        {
            BindCombo();
        }

        private void BindCombo()
        {
            #region 店舗

            DataTable dtstore = new DataTable();

            dtstore.Columns.Add("storeID");
            dtstore.Columns.Add("storeName");

            dtstore.Rows.Add("-1", "");
            dtstore.Rows.Add("1", "本社Web");
            dtstore.Rows.Add("2", "豊中店");
            dtstore.Rows.Add("3", "石橋店");
            dtstore.Rows.Add("4", "三宮店");
            dtstore.Rows.Add("5", "江坂店");

            dtstore.AcceptChanges();
            combostore.SelectedIndex = -1;

            combostore.DisplayMember = "storeName";
            combostore.ValueMember = "storeID";

            combostore.DataSource = dtstore;
            #endregion
        }

       
    }
}
