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
    public partial class frmBrand_List : Form
    {
        public string brandName, brandCD = string.Empty;
        M_Brand_Entity m_Brand_Data = new M_Brand_Entity();
        Brand_List_BL Brand_List_BL = new Brand_List_BL();

        /// <summary>
        /// default
        /// </summary>
        public frmBrand_List()
        {
            InitializeComponent();
            //SetDesignFunction();
        }

        /// <summary>
        /// form load
        /// </summary>
        private void frmBrand_List_Load(object sender, EventArgs e)
        {
            dgvBrand.DisabledColumn("colNo");
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
            m_Brand_Data = new M_Brand_Entity();
            m_Brand_Data.BrandName = txtBrandName.Text;
            DataTable dt_Brand = new DataTable();
            dt_Brand = Brand_List_BL.Brand_Select(m_Brand_Data);
            dgvBrand.DataSource = dt_Brand;
            if (dt_Brand.Rows.Count > 0)
            {
                dgvBrand.Rows[0].Selected = true;
            }
        }

        /// <summary>
        ///  Bind data to Main Form
        /// </summary>
        private void F12()
        {
            if (dgvBrand.Rows.Count > 0)
            {
                brandCD = dgvBrand.Rows[dgvBrand.CurrentRow.Index].Cells["colBrandCD"].Value.ToString();
                brandName = dgvBrand.Rows[dgvBrand.CurrentRow.Index].Cells["colBrandName"].Value.ToString();
            }
            this.Close();
        }

        /// <summary>
        /// Bind data to main form when double click Grid cell
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvBrand_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                var senderGrid = (DataGridView)sender;
                if (!String.IsNullOrWhiteSpace(senderGrid.Rows[e.RowIndex].Cells["colBrandCD"].Value.ToString()))
                {
                    brandCD = senderGrid.Rows[e.RowIndex].Cells["colBrandCD"].Value.ToString();
                }

                if (!String.IsNullOrWhiteSpace(senderGrid.Rows[e.RowIndex].Cells["colBrandName"].Value.ToString()))
                {
                    brandName = senderGrid.Rows[e.RowIndex].Cells["colBrandName"].Value.ToString();
                }
                this.Close();
            }
        }

        /// <summary>
        /// handle f1 to f12 click event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmBrand_List_KeyDown(object sender, KeyEventArgs e)
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
