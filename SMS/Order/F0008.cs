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
    public partial class F0008 : SMS_Base
    {
        Login_Info loginInfo;

        public F0008()
        {
            InitializeComponent();
        }

        public F0008(Login_Info loginInfo)
            : base(loginInfo)
        {
            InitializeComponent();
            this.loginInfo = loginInfo;
            SetDesignerFunction();
        }

        private void SetDesignerFunction()
        {
            lblMode.Visible = false;

            dgv1.DisabledColumn("colNo,colSupplier,colOrderCount,colCaptureNo,colErrorNo");

            FunctionButtonDisabled(2);
            FunctionButtonDisabled(3);
            FunctionButtonDisabled(4);
            FunctionButtonDisabled(6);
            FunctionButtonDisabled(7);
            FunctionButtonDisabled(8);
        }

        private void F0008_Load(object sender, EventArgs e)
        {
            FormName = "EDI回答納期登録";
            //BindGrid();
        }

        private void BindGrid()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("No");
            dt.Columns.Add("Supplier");
            dt.Columns.Add("OrderCount");
            dt.Columns.Add("CaptureNo");
            dt.Columns.Add("ErrorNo");

            dt.Rows.Add("1", "XXXXXXXXXX ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "9,999", "9,999", "9,999");
            dt.Rows.Add("2", "XXXXXXXXXX ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "9,999", "9,999", "9,999");
            dt.Rows.Add("3", "XXXXXXXXXX ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "9,999", "9,999", "9,999");
            dt.Rows.Add("4", "XXXXXXXXXX ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "9,999", "9,999", "9,999");
            dt.Rows.Add("5", "XXXXXXXXXX ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "9,999", "9,999", "9,999");
            dt.Rows.Add("6", "XXXXXXXXXX ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "9,999", "9,999", "9,999");
            dt.Rows.Add("7", "XXXXXXXXXX ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "9,999", "9,999", "9,999");
            dt.Rows.Add("8", "XXXXXXXXXX ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "9,999", "9,999", "9,999");
            dt.Rows.Add("9", "XXXXXXXXXX ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "9,999", "9,999", "9,999");
            dt.Rows.Add("10", "XXXXXXXXXX ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "9,999", "9,999", "9,999");
            dt.Rows.Add("11", "XXXXXXXXXX ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "9,999", "9,999", "9,999");
            dt.Rows.Add("12", "XXXXXXXXXX ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "9,999", "9,999", "9,999");
            dt.Rows.Add("13", "XXXXXXXXXX ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "9,999", "9,999", "9,999");
            dt.Rows.Add("14", "XXXXXXXXXX ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "9,999", "9,999", "9,999");
            dt.Rows.Add("15", "XXXXXXXXXX ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "9,999", "9,999", "9,999");
            dt.AcceptChanges();

            dgv1.AutoGenerateColumns = false;
            dgv1.DataSource = dt;

        }

        private void btnCheckAll_Click(object sender, EventArgs e)
        {
            //foreach (DataGridViewRow row in smS_GridView1.Rows)
            //{
            //    DataGridViewCheckBoxCell chk = (DataGridViewCheckBoxCell)row.Cells[1];
            //    chk.TrueValue = true;
            //    chk.FalseValue = false;
            //    if (chk.Value == chk.TrueValue)
            //    {
            //        chk.Value = chk.TrueValue;
            //    }
            //    else
            //    {
            //        chk.Value = chk.TrueValue;
            //    }
            //}

            for (int i = 0; i < dgv1.Rows.Count; i++)
            {

                dgv1[1, i].Value = true;//check every row in the forth column
                
            }
        }

        private void btnUnCheckAll_Click(object sender, EventArgs e)
        {
            //foreach (DataGridViewRow row in smS_GridView1.Rows)
            //{

            //    DataGridViewCheckBoxCell chk = (DataGridViewCheckBoxCell)row.Cells[1];
            //    chk.TrueValue = true;
            //    chk.FalseValue = false;
            //    if (chk.Value == chk.TrueValue)
            //    {
            //        chk.Value = chk.FalseValue;
            //    }
            //    else
            //    {
            //        chk.Value = chk.FalseValue;
            //    }
            //}

            for (int i = 0; i < dgv1.Rows.Count; i++)
            {

                dgv1[1, i].Value = false;//check every row in the forth column

            }
        }

        public override void functionClick(int btnIndex)
        {
            switch (btnIndex)
            {
                case 10:
                    F10();
                    break;
            }
        }

        private void F10()
        {
            frmF0009 frmf9 = new frmF0009();
            frmf9.Show();
        }
    }
}
