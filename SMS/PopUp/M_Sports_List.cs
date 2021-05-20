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
    public partial class frmM_Sports_List : Form
    {
        DataTable dtkyogi;
        m_sports_BL msbl;
        m_sports_Entity mse;
        public string nc_sports;
        public string vm_sports;

        /// <summary>
        /// default
        /// </summary>
        public frmM_Sports_List()
        {
            InitializeComponent();
            this.MaximizeBox = false;
            SetDesignerFunction();
        }
        
        /// <summary>
        /// Form Load
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void M_Sports_List_Load(object sender, EventArgs e)
        {
            this.KeyPreview = true;
            msbl = new m_sports_BL();
            BindGrid();
        }

        /// <summary>
        /// Design Setting 
        /// </summary>
        private void SetDesignerFunction()
        {
            gvkyogi.CellDoubleClick += gvkyogi_CellDoubleClick;
            btnF1.Click += btnF1_Click;
            this.StartPosition = FormStartPosition.CenterParent;
        }

        /// <summary>
        /// Data Binding for DataGridView
        /// </summary>
        private void BindGrid()
        {
            mse = new m_sports_Entity();
            mse.vm_sports = txtvmsports.Text;
            dtkyogi = msbl.m_sports_Select(mse);
            gvkyogi.DataSource = dtkyogi;
            if (dtkyogi.Rows.Count > 0)
            {
                gvkyogi.Rows[0].Selected = true;
            }
        }

        /// <summary>
        /// handle f1 to f12 click event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void M_Sports_List_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F1)
                this.Close();
            else if (e.KeyCode == Keys.F11)
                BindGrid();
            else if (e.KeyCode == Keys.F12)
                F12();
        }

        /// <summary>
        /// Bind Data to Main Form When Click on F12 Button
        /// </summary>
        private void F12()
        {
            if (gvkyogi.Rows.Count > 0)
            {
                nc_sports = gvkyogi.Rows[gvkyogi.CurrentRow.Index].Cells["colSportsCD"].Value.ToString();
                vm_sports = gvkyogi.Rows[gvkyogi.CurrentRow.Index].Cells["colSportsName"].Value.ToString();
            }
            this.Close();
        }

        #region Button Click
        private void btnF12_Click(object sender, EventArgs e)
        {
            F12();
        }

        private void btnF11_Click(object sender, EventArgs e)
        {
            BindGrid();
        }

        private void btnF1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        #endregion

        /// <summary>
        /// Bind data to main form when double click Grid cell
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void gvkyogi_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                var senderGrid = (DataGridView)sender;
                if (!String.IsNullOrWhiteSpace(senderGrid.Rows[e.RowIndex].Cells["colSportsCD"].Value.ToString()))
                {
                    nc_sports = senderGrid.Rows[e.RowIndex].Cells["colSportsCD"].Value.ToString();
                }　　

                if (!String.IsNullOrWhiteSpace(senderGrid.Rows[e.RowIndex].Cells["colSportsName"].Value.ToString()))
                {
                    vm_sports = senderGrid.Rows[e.RowIndex].Cells["colSportsName"].Value.ToString();
                }
                this.Close();
            }
        }
    }
}
