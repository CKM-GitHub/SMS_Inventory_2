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
using System.IO;

namespace SMS.Order
{
    public partial class F0003 : SMS_Base
    {
        Login_Info loginInfo;

        public F0003()
        {
            InitializeComponent();            
        }

        public F0003(Login_Info loginInfo)
            : base(loginInfo)
        {
            this.loginInfo = loginInfo;
            InitializeComponent();
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

            dgvF0003.DisabledColumn("colNo,colNayosemachi,colMallStore,colTorikomikensu,colTorikomiNichiji,colTorikomiOpreator,colHoryuTaisho,colErrorCount");

            checkBox18.Checked = true;
        }

        private void BindCombo()
        {
            DataTable dtMall = new DataTable();
            dtMall.Columns.Add("MallID");
            dtMall.Columns.Add("MallName");

            dtMall.Rows.Add("-1", "");
            dtMall.Rows.Add("0", "Yahoo Racket Plaza");
            dtMall.Rows.Add("1", "Yahoo Sports Plaza");
            dtMall.Rows.Add("2", "Yahoo LUCK PIECE");
            dtMall.Rows.Add("3", "Yahoo Baseball Plaza");
            dtMall.Rows.Add("4", "楽天 Racket Plaza");
            dtMall.Rows.Add("5", "楽天 Sports Plaza");
            dtMall.Rows.Add("6", "楽天 LUCK PIECE");
            dtMall.Rows.Add("7", "楽天 Baseball Plaza");
            dtMall.Rows.Add("8", "Amazon Racket Plaza");
            dtMall.Rows.Add("9", "ポンパレ Racket Plaza");
            dtMall.AcceptChanges();
            comboMall.SelectedIndex = -1;

            this.comboMall.DisplayMember = "MallName";
            this.comboMall.ValueMember = "MallID";
            this.comboMall.DataSource = dtMall;
        }

        private void F0003_Load(object sender, EventArgs e)
        {
            SetDesignerFunction();
            FormName = "WEB受注取込(CSV)";
            BindGrid();
            BindCombo();
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

            string[] row1 = new string[] { "1", "2018/8/20 12:30", "RacketPlaza", "自動", "199", "16", "20", "12" };
            string[] row2 = new string[] { "2", "2018/9/11 11:04", "自社", "自動", "60", "10", "14", "6" };
            string[] row3 = new string[] { "3", "2018/9/20 17:30", "Yahoo　SportsPlaza", "自動", "109", "10", "40", "20" };
            string[] row4 = new string[] { "4", "2018/9/20 12:34", "自社", "自動", "87", "34", "11", "6" };
            string[] row5 = new string[] { "5", "2018/10/5 13:48", "Wowma", "自動", "104", "4", "14", "21" };
            string[] row6 = new string[] { "6", "2018/10/10 10:44", "Yahoo　RacketPlaza", "自動", "231", "12", "6", "3" };
            string[] row7 = new string[] { "7", "2018/10/20 12:34", "楽天　 RacketPlaza", "自動", "120", "5", "63", "11" };
            string[] row8 = new string[] { "8", "2018/10/17 8:09", "Yahoo　RacketPlaza", "自動", "160", "6", "2", "14" };
            string[] row9 = new string[] { "9", "2018/10/25 5:05", "Yahoo　RacketPlaza", "自動", "200", "65", "8", "10" };
            string[] row10 = new string[] { "10", "2018/11/6 1:00", "Yahoo　LUCK PIECE", "自動", "140", "9", "16", "3" };
            string[] row11 = new string[] { "11", "2018/11/10 10:50", "楽天　 RacketPlaza", "自動", "122", "18", "14", "33" };
            string[] row12 = new string[] { "12", "2018/11/15 13:30", "ポンパレ RacketPlaza", "自動", "133", "21", "5", "7" };
            string[] row13 = new string[] { "13", "2018/11/20 20:11", "楽天　 RacketPlaza", "自動", "335", "6", "14", "18" };
            string[] row14 = new string[] { "14", "2018/12/1 1:00", "ポンパレ RacketPlaza", "自動", "125", "67", "24", "6" };
            string[] row15 = new string[] { "15", "2018/12/5 17:30", "楽天　 RacketPlaza ", "自動", "314", "42", "31", "16" };

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

            dgvF0003.AutoGenerateColumns = false;
            dgvF0003.DataSource = dt;
        }

        private void dgvF0003_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if ((sender as DataGridView).CurrentCell is DataGridViewCheckBoxCell)
            {
                if ((Convert.ToBoolean(dgvF0003.Rows[e.RowIndex].Cells[e.ColumnIndex].EditedFormattedValue) == true))
                {
                    foreach (DataGridViewRow row1 in dgvF0003.Rows)
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

                    dgvF0003.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = true;
                    dgvF0003.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Yellow;
                }
                else
                {
                    dgvF0003.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.White;
                    dgvF0003.ClearSelection();
                }

                dgvF0003.EndEdit();
            }
        }

        private void btnShowDialog_Click(object sender, EventArgs e)
        {
            string folderPath = "C:\\Csv\\";
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
                System.Diagnostics.Process.Start(folderPath);
            }
            else
            {
                System.Diagnostics.Process.Start(folderPath);
            }
        }
    }
}
