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
    public partial class frmF0002 : SMS_Base
    {
        Login_Info loginInfo;

        public frmF0002()
        {
            InitializeComponent();
        }

        public frmF0002(Login_Info loginInfo)
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
            FunctionButtonDisabled(9);
            FunctionButtonDisabled(10);

            checkBox1.Checked = true;
            checkBox2.Checked = true;
            checkBox4.Checked = true;
            checkBox10.Checked = true;
            checkBox13.Checked = true;
            //chkName.Checked = true;
            //chkOrderData.Checked = true;
            //chkAutoReserve.Checked = true;
            //chkAutoOrder.Checked = true;
            chkNameIdentification.Checked=true;

            dgvF0002.DisabledColumn("colNo,colNayosemachi,colMallStore,colTorikomikensu,colTorikomiNichiji,colTorikomiOpreator,colHoryuTaisho,colErrorCount");
        }

        private void frmF0002_Load(object sender, EventArgs e)
        {
            BindGrid();
            FormName = "WEB受注取込";
        }

        private void BindGrid()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("No");
            dt.Columns.Add("TorikomiNichiji");
            dt.Columns.Add("MallStore");
            dt.Columns.Add("TorikomiOpreator");
            dt.Columns.Add("Torikomikensu");
            dt.Columns.Add("ErrorCount");
            dt.Columns.Add("HoryuTaisho");
            dt.Columns.Add("Nayosemachi");

            string[] row1 = new string[] { "1", "2018/8/20 12:30", "RacketPlaza", "自動", "141", "5", "7","7" };
            string[] row2 = new string[] { "2", "2018/9/11 11:04", "自社", "自動", "162", "5", "4", "1" };
            string[] row3 = new string[] { "3", "2018/9/20 17:30", "Yahoo　SportsPlaza", "自動", "107", "7", "10", "8" };
            string[] row4 = new string[] { "4", "2018/9/20 12:34", "自社", "自動", "164", "4", "4", "7" };
            string[] row5 = new string[] { "5", "2018/10/5 13:48", "Wowma", "自動", "78", "6", "14", "3" };
            string[] row6 = new string[] { "6", "2018/10/10 10:44", "Yahoo　RacketPlaza", "自動", "64", "7", "3", "8" };
            string[] row7 = new string[] { "7", "2018/10/20 12:34", "楽天　 RacketPlaza", "自動", "213", "20", "11", "6" };
            string[] row8 = new string[] { "8", "2018/10/17 8:09", "Yahoo　RacketPlaza", "自動", "102", "10", "3", "4" };
            string[] row9 = new string[] { "9", "2018/10/25 5:05", "Yahoo　RacketPlaza", "自動", "74", "10", "6", "4" };
            string[] row10 = new string[] { "10", "2018/11/6 1:00", "Yahoo　LUCK PIECE", "自動", "117", "10", "5", "17" };
            string[] row11 = new string[] { "11", "2018/11/10 10:50", "楽天　 RacketPlaza", "自動", "67", "4", "4", "3" };
            string[] row12 = new string[] { "12", "2018/11/15 13:30", "ポンパレ RacketPlaza", "自動", "81", "10", "5", "3" };
            string[] row13 = new string[] { "13", "2018/11/20 20:11", "楽天　 RacketPlaza", "自動", "204", "10", "4", "11" };
            string[] row14 = new string[] { "14", "2018/12/1 1:00", "ポンパレ RacketPlaza", "自動", "153", "2", "5", "6" };
            string[] row15 = new string[] { "15", "2018/12/5 17:30", "楽天　 RacketPlaza ", "自動", "98", "6", "13", "2" };

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

            dgvF0002.AutoGenerateColumns = false;
            dgvF0002.DataSource = dt;
        }

        private void btnAllCheck_Click(object sender, EventArgs e)
        {
            chkName.Checked = true;
            chkOrderData.Checked = true;
            chkAutoOrder.Checked = true;
            chkAutoReserve.Checked = true;
            chkNameIdentification.Checked = true;
        }

        private void btnAllCancel_Click(object sender, EventArgs e)
        {
            chkName.Checked = false;
            chkOrderData.Checked = false;
            chkAutoOrder.Checked = false;
            chkAutoReserve.Checked = false;
            chkNameIdentification.Checked = false;
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            lblKakuteiMode.Text = "処理実行中";
            lblKakuteiMode.BackColor = Color.FromArgb(0, 176, 240);
        }

        private void btnStop_Click(object sender, EventArgs e)
        {
            lblKakuteiMode.Text = "処理停止中";
            lblKakuteiMode.BackColor = Color.FromArgb(255, 102, 255);
        }

        private void dgvF0002_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if ((sender as DataGridView).CurrentCell is DataGridViewCheckBoxCell)
            {
                if ((Convert.ToBoolean(dgvF0002.Rows[e.RowIndex].Cells[e.ColumnIndex].EditedFormattedValue) == true))
                {
                    foreach (DataGridViewRow row1 in dgvF0002.Rows)
                    {
                        DataGridViewCheckBoxCell chk1 = row1.Cells["chk1"] as DataGridViewCheckBoxCell;
                        DataGridViewCheckBoxCell chk2 = row1.Cells["chk2"] as DataGridViewCheckBoxCell;

                        chk1.Value = chk1.FalseValue;
                        chk2.Value = chk2.FalseValue;

                        if (row1.Index % 2 == 0)
                        {
                            row1.DefaultCellStyle.BackColor = Color.White;
                        }
                        else
                        {
                            row1.DefaultCellStyle.BackColor = Color.FromArgb(221, 235, 247); ;
                        }
                    }

                    dgvF0002.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = true;
                    dgvF0002.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Yellow;
                }
                else
                {
                    dgvF0002.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.White;
                    dgvF0002.ClearSelection();
                }

                dgvF0002.EndEdit();
            }
        }   
    }    
}
 