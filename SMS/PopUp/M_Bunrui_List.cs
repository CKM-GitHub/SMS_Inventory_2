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
    public partial class frmM_Bunrui_List : Form
    {
        DataTable dtBunrui;
        m_bunrui_Entity mbe;
        m_bunrui_BL mbbl;
        public string nc_bunrui;
        public string vm_bunrui;

        /// <summary>
        /// default
        /// </summary>
        public frmM_Bunrui_List()
        {
            InitializeComponent();
        }

        /// <summary>
        /// form load
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmM_Bunrui_List_Load(object sender, EventArgs e)
        {
            this.KeyPreview = true;
            mbbl = new m_bunrui_BL();
            BindGrid();
        }
        
        #region Button Click
        /// <summary>
        /// Close Form
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnF1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// Search
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnF11_Click(object sender, EventArgs e)
        {
            BindGrid();
        }

        /// <summary>
        /// Bind data to Main Form
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnF12_Click(object sender, EventArgs e)
        {
            F12();
        }

        #endregion

        /// <summary>
        /// bind grid data
        /// </summary>
        private void BindGrid()
        {
            mbe = new m_bunrui_Entity();
            mbe.vm_bunrui = txtvmbunrui.Text;
            dtBunrui = mbbl.m_bunrui_Select(mbe);
            dgvBunrui.DataSource = dtBunrui;
            if (dtBunrui.Rows.Count > 0)
            {
                dgvBunrui.Rows[0].Selected = true;
            }
        }

        /// <summary>
        /// Bind data to Main Form
        /// </summary>
        private void F12()
        {

            nc_bunrui = dgvBunrui.Rows[dgvBunrui.CurrentRow.Index].Cells["colCategoryCD"].Value.ToString();
            vm_bunrui = dgvBunrui.Rows[dgvBunrui.CurrentRow.Index].Cells["colCategoryName"].Value.ToString();

            this.Close();
        }

        /// <summary>
        /// Bind data to main form when double click Grid cell
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvBunrui_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                var senderGrid = (DataGridView)sender;
                if (!String.IsNullOrWhiteSpace(senderGrid.Rows[e.RowIndex].Cells["colCategoryCD"].Value.ToString()))
                {
                    nc_bunrui = senderGrid.Rows[e.RowIndex].Cells["colCategoryCD"].Value.ToString();
                }

                if (!String.IsNullOrWhiteSpace(senderGrid.Rows[e.RowIndex].Cells["colCategoryName"].Value.ToString()))
                {
                    vm_bunrui = senderGrid.Rows[e.RowIndex].Cells["colCategoryName"].Value.ToString();
                }
                this.Close();
            }
        }

        /// <summary>
        /// handle for Searching with function key
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmM_Bunrui_List_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F1)
                this.Close();
            else if (e.KeyCode == Keys.F11)
                BindGrid();
            else if (e.KeyCode == Keys.F12)
                F12();
        }

        
    }
}
