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
    public partial class ComboBoxTest : Form
    {
        public ComboBoxTest()
        {
            InitializeComponent();
            DataTable dt = new DataTable();

            dt.Columns.Add("warehouseID");
            dt.Columns.Add("warehouseName");


            dt.Rows.Add("-1", "");
            dt.Rows.Add("1", "本社Web倉庫");
            dt.Rows.Add("2", "豊中店倉庫");
            dt.Rows.Add("3", "石橋店倉庫");
            dt.Rows.Add("4", "三宮店倉庫");
            dt.Rows.Add("5", "江坂店倉庫");

            dt.AcceptChanges();
            ckM_ComboBox1.SelectedIndex = -1;

            ckM_ComboBox1.DisplayMember = "warehouseName";
            ckM_ComboBox1.ValueMember = "warehouseID";

            this.ckM_ComboBox1.DataSource = dt;
        }
    }
}
