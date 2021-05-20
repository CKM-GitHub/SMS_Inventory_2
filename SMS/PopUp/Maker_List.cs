using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SMS_BL;
using SMS_Entity;

namespace SMS.PopUp
{
    public partial class frmMaker_List : Form
    {
        public string MakerCD, MakerName = string.Empty;
        M_SHIIRESAKI_Entity M_Shiiresaki_Data = new M_SHIIRESAKI_Entity();
        Maker_List_BL MakerList_BL = new Maker_List_BL();

        /// <summary>
        /// default
        /// </summary>
        public frmMaker_List() 
        {
            InitializeComponent();
            this.MaximizeBox = false;
            dgvMakerSearch.DisabledColumn("colNo");
        }

        /// <summary>
        /// form load
        /// </summary>
        private void frmMaker_List_Load(object sender, EventArgs e)
        {
            BindGrid();
        }

        /// <summary>
        /// bind grid data
        /// </summary>
        private void BindGrid()
        {
            M_Shiiresaki_Data = new M_SHIIRESAKI_Entity();
            M_Shiiresaki_Data.MakerName = txtMakerName.Text;
            DataTable dt_MakerSearch = new DataTable();
            dt_MakerSearch = MakerList_BL.Maker_Select(M_Shiiresaki_Data);
            dgvMakerSearch.DataSource = dt_MakerSearch;
            if (dt_MakerSearch.Rows.Count > 0)
            {
                dgvMakerSearch.Rows[0].Selected = true;
            }
        }

        private void F12()
        {
            if (dgvMakerSearch.Rows.Count > 0)
            {
                MakerCD = dgvMakerSearch.Rows[dgvMakerSearch.CurrentRow.Index].Cells["colMakerCD"].Value.ToString();
                MakerName = dgvMakerSearch.Rows[dgvMakerSearch.CurrentRow.Index].Cells["colMakerName"].Value.ToString();
            }
            this.Close();
        }

        #region Button Click
        /// <summary>
        /// form close
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnF1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnF11_Click(object sender, EventArgs e)
        {
            BindGrid();
        }

        private void btnF12_Click(object sender, EventArgs e)
        {
            F12();
        }
        #endregion

        /// <summary>
        /// Bind data to main form when double click Grid cell
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvMakerSearch_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                var senderGrid = (DataGridView)sender;
                if (!String.IsNullOrWhiteSpace(senderGrid.Rows[e.RowIndex].Cells["colMakerCD"].Value.ToString()))
                {
                    MakerCD = senderGrid.Rows[e.RowIndex].Cells["colMakerCD"].Value.ToString();
                }

                if (!String.IsNullOrWhiteSpace(senderGrid.Rows[e.RowIndex].Cells["colMakerName"].Value.ToString()))
                {
                    MakerName = senderGrid.Rows[e.RowIndex].Cells["colMakerName"].Value.ToString();
                }
                this.Close();
            }
        }

        /// <summary>
        /// handle f1 to f12 click event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmMaker_List_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F1)
                this.Close();
            else if (e.KeyCode == Keys.F12)
                F12();
            else if (e.KeyCode == Keys.F11)
                BindGrid();
        }
    }
}
