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
    public partial class frmPattern_List : Form
    {
        public string PatternCD, PatternName = string.Empty;
        M_MakerZaiko_H_Entity maker_H_Data = new M_MakerZaiko_H_Entity();
        Pattern_List_BL Pattern_BL = new Pattern_List_BL();

       /// <summary>
       /// default
       /// </summary>
        public frmPattern_List()
        {
            InitializeComponent();
            SetDesignFunction();
            this.MaximizeBox = false;
        }
        
        /// <summary>
        /// Design Setting
        /// </summary>
        private void SetDesignFunction()
        {
            dgvMKZH.DisabledColumn("colNo");
        }

        /// <summary>
        /// Form Load
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmPattern_List_Load(object sender, EventArgs e)
        {
            BindGrid();
        }

        /// <summary>
        /// Bind Data to DataGridView
        /// </summary>
        private void BindGrid()
        {
            maker_H_Data = new M_MakerZaiko_H_Entity();
            maker_H_Data.PatternName = txtPatternName.Text;
            DataTable dt_MZKH = new DataTable();
            dt_MZKH = Pattern_BL.M_MakerZaiko_H_Select(maker_H_Data);
            dgvMKZH.DataSource = dt_MZKH;
            if (dt_MZKH.Rows.Count > 0)
            {
                dgvMKZH.Rows[0].Selected = true;
            }
        }

        /// <summary>
        /// Bind Data to Main Form When Click on F12 Button
        /// </summary>
        private void F12()
        {
            if (dgvMKZH.Rows.Count > 0)
            {
                PatternCD = dgvMKZH.Rows[dgvMKZH.CurrentRow.Index].Cells["colPatternCD"].Value.ToString();
                PatternName = dgvMKZH.Rows[dgvMKZH.CurrentRow.Index].Cells["colPatternName"].Value.ToString();
            }
            this.Close();
        }

        #region Button Click
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
        /// Bind datat to Main Form When double click DataGrid Cell 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvMKZH_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                var senderGrid = (DataGridView)sender;
                if (!String.IsNullOrWhiteSpace(senderGrid.Rows[e.RowIndex].Cells["colPatternCD"].Value.ToString()))
                {
                    PatternCD = senderGrid.Rows[e.RowIndex].Cells["colPatternCD"].Value.ToString();

                }

                if (!String.IsNullOrWhiteSpace(senderGrid.Rows[e.RowIndex].Cells["colPatternName"].Value.ToString()))
                {
                    PatternName = senderGrid.Rows[e.RowIndex].Cells["colPatternName"].Value.ToString();
                }
                this.Close();
            }
        }

        /// <summary>
        /// Keydown Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmPattern_List_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F1)
                this.Close();
            if (e.KeyCode == Keys.F11)
                BindGrid();
            else if (e.KeyCode == Keys.F12)
                F12();
        }

    }
}
