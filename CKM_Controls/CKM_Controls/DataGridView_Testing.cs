using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CKM_Controls
{
    public partial class DataGridView_Testing : Form
    {
        public DataGridView_Testing()
        {
            InitializeComponent();
        }

        private void DataGridView_Testing_Load(object sender, EventArgs e)
        {
            ckM_DataGridView2.DisabledColumn("Column1,Column5");

            DataTable dt = new DataTable();
            dt.Columns.Add("Column1");
            dt.Columns.Add("Column2");
            dt.Columns.Add("Column5");
            dt.Columns.Add("Column6");
            dt.Columns.Add("Column7");
            dt.Columns.Add("Column8");
            for (int i = 0; i <= 4; i++)
            {
                dt.Rows.Add();
                dt.Rows[i]["Column1"] = "Cat";
                dt.Rows[i]["Column2"] = i;
                dt.Rows[i]["Column5"] = i + 1;
                dt.Rows[i]["Column6"] = i + 2;
                dt.Rows[i]["Column7"] = i + 3;
                dt.Rows[i]["Column8"] = i + 3;
            }               
            ckM_DataGridView2.DataSource = dt;
        }
    }
}
