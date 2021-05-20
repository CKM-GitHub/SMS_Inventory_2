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
    public partial class F0007 : SMS_Base
    {
        Login_Info loginInfo;

        public F0007()
        {
            InitializeComponent();
        }

        public F0007(Login_Info loginInfo)
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
            FunctionButtonDisabled(6);
            FunctionButtonDisabled(7);
            FunctionButtonDisabled(8);
            FunctionButtonDisabled(12);

            dgvF0007.DisabledColumn("colNo,colHatchuNumber,colShiiresaki,colHatchuSu");
        }

        private void BindCombo()
        {
            DataTable dtTarget = new DataTable();
            dtTarget.Columns.Add("TargetID");
            dtTarget.Columns.Add("TargetName");

            dtTarget.Rows.Add("-1", "");
            dtTarget.Rows.Add("0", "Web店");
            dtTarget.Rows.Add("1", "豊中店");
            dtTarget.Rows.Add("2", "石橋店");
            dtTarget.Rows.Add("3", "三宮店");
            dtTarget.Rows.Add("4", "江坂店");
            dtTarget.Rows.Add("5", "楽天ラケットプラザ");
            dtTarget.Rows.Add("6", "Yahooラケットプラザ");

            dtTarget.AcceptChanges();
            comboTarget.SelectedIndex = -1;

            this.comboTarget.DisplayMember = "TargetName";
            this.comboTarget.ValueMember = "TargetID";
            this.comboTarget.DataSource = dtTarget;
        }

        private void F0007_Load(object sender, EventArgs e)
        {
            BindCombo();
            //BindGrid();
            FormName = "発注EDIデータ作成";
        }

        private void BindGrid()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("No");
            dt.Columns.Add("HatchuNumber");
            dt.Columns.Add("Shiiresaki");
            dt.Columns.Add("HatchuSu");

            string[] row1 = new string[] { "1", "XXXXXXXXX11", "XXXXXXXXXX ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "9,999" };
            string[] row2 = new string[] { "2", "XXXXXXXXX11", "XXXXXXXXXX ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "9,999" };
            string[] row3 = new string[] { "3", "XXXXXXXXX11", "XXXXXXXXXX ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "9,999" };
            string[] row4 = new string[] { "4", "XXXXXXXXX11", "XXXXXXXXXX ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "9,999" };
            string[] row5 = new string[] { "5", "XXXXXXXXX11", "XXXXXXXXXX ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "9,999" };
            string[] row6 = new string[] { "6", "XXXXXXXXX11", "XXXXXXXXXX ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "9,999" };
            string[] row7 = new string[] { "7", "XXXXXXXXX11", "XXXXXXXXXX ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "9,999" };
            string[] row8 = new string[] { "8", "XXXXXXXXX11", "XXXXXXXXXX ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "9,999" };
            string[] row9 = new string[] { "9", "XXXXXXXXX11", "XXXXXXXXXX ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "9,999" };
            string[] row10 = new string[] { "10", "XXXXXXXXX11", "XXXXXXXXXX ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "9,999" };
            string[] row11 = new string[] { "11", "XXXXXXXXX11", "XXXXXXXXXX ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "9,999" };
            string[] row12 = new string[] { "12", "XXXXXXXXX11", "XXXXXXXXXX ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "9,999" };
            string[] row13 = new string[] { "13", "XXXXXXXXX11", "XXXXXXXXXX ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "9,999" };
            string[] row14 = new string[] { "14", "XXXXXXXXX11", "XXXXXXXXXX ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "9,999" };
            string[] row15 = new string[] { "15", "XXXXXXXXX11", "XXXXXXXXXX ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "9,999" };
            string[] row16 = new string[] { "16", "XXXXXXXXX11", "XXXXXXXXXX ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "9,999" };
            string[] row17 = new string[] { "17", "XXXXXXXXX11", "XXXXXXXXXX ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "9,999" };
            string[] row18 = new string[] { "18", "XXXXXXXXX11", "XXXXXXXXXX ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "9,999" };
            string[] row19 = new string[] { "19", "XXXXXXXXX11", "XXXXXXXXXX ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "9,999" };
            string[] row20 = new string[] { "20", "XXXXXXXXX11", "XXXXXXXXXX ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", "9,999" };

            dt.Rows.Add(row1);
            dt.Rows.Add(row2);
            dt.Rows.Add(row3);
            dt.Rows.Add(row4);
            dt.Rows.Add(row5);
            dt.Rows.Add(row6);
            dt.Rows.Add(row7);
            dt.Rows.Add(row8);
            dt.Rows.Add(row9);
            dt.Rows.Add(row10);
            dt.Rows.Add(row11);
            dt.Rows.Add(row12);
            dt.Rows.Add(row13);
            dt.Rows.Add(row14);
            dt.Rows.Add(row15);
            dt.Rows.Add(row16);
            dt.Rows.Add(row17);
            dt.Rows.Add(row18);
            dt.Rows.Add(row19);
            dt.Rows.Add(row20);

            dgvF0007.AutoGenerateColumns = false;
            dgvF0007.DataSource = dt;
        }

        private void btnCheckall_Click(object sender, EventArgs e)
        {

            //foreach (DataGridViewRow row in dgvF0007.Rows)
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
            for (int i = 0; i < dgvF0007.Rows.Count; i++)
            {

                dgvF0007[1, i].Value = true;//check every row in the forth column

            }
        }

        private void btnUnCheckall_Click(object sender, EventArgs e)
        {
            //foreach (DataGridViewRow row in dgvF0007.Rows)
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
            for (int i = 0; i < dgvF0007.Rows.Count; i++)
            {

                dgvF0007[1, i].Value = false;//check every row in the forth column

            }
            
        }
    }
}
