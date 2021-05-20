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
    public partial class F0004 : SMS_Base
    {
        Login_Info loginInfo;

        public F0004()
        {
            InitializeComponent();
        }

        public F0004(Login_Info loginInfo)
            : base(loginInfo)
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
            FunctionButtonDisabled(5);
            FunctionButtonDisabled(6);
            FunctionButtonDisabled(7);
            FunctionButtonDisabled(8);
            FunctionButtonDisabled(9);
            FunctionButtonDisabled(10);
            FunctionButtonDisabled(11);
        }

        private void F0004_Load(object sender, EventArgs e)
        {
            BindCombo();
            FormName = "各種処理起動";
        }

        private void BindCombo()
        {
            DataTable dt = new DataTable();

            dt.Columns.Add("processID");
            dt.Columns.Add("processName");


            dt.Rows.Add("-1", "");

            dt.Rows.Add("1", "名寄せ処理(要注意顧客)");
            dt.Rows.Add("2", "受注データチェック処理");
            dt.Rows.Add("3", "自動引当処理");
            dt.Rows.Add("4", "自動発注判断");
            dt.Rows.Add("5", "名寄せ処理(全顧客)");
            dt.AcceptChanges();
            cboprocessing.SelectedIndex = -1;

            cboprocessing.DisplayMember = "processName";
            cboprocessing.ValueMember = "processID";

            this.cboprocessing.DataSource = dt;
        }

        //private void panelBody_Paint(object sender, PaintEventArgs e)
        //{

        //}
    }
}
