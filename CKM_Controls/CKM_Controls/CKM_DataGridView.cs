using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using System.Collections;


namespace CKM_Controls
{
    public class CKM_DataGridView : DataGridView
    {
        public bool IsCellValid { get; set; }

        string DisablecolName = string.Empty;

        string EnablecolName = string.Empty;

        public CKM_DataGridView()
        {
            this.DefaultCellStyle.Font = new Font("MS Gothic", 9, FontStyle.Regular);

            this.EnableHeadersVisualStyles = false;
            this.ColumnHeadersDefaultCellStyle.Font = new Font(this.Font, FontStyle.Bold);
            this.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(191, 191, 191);
            this.BackgroundColor = this.GridColor = Color.FromArgb(198, 224, 180);

            this.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(221, 235, 247);
            this.RowHeadersVisible = false;
            this.AllowUserToAddRows = false;

            this.AutoGenerateColumns = false;
            this.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.AllowUserToDeleteRows = false;
            this.AllowUserToResizeRows = false;
            this.DisabledSorting();
          // this.DisableHeaderRowResize();
        }

        public void DisableHeaderRowResize()
        {
            this.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;

            for (int i = 0; i < this.Columns.Count; i ++ )
            {
                this.Columns[i].Resizable = DataGridViewTriState.False;
            }


        }

        public void DisabledSorting()
        {
            foreach (DataGridViewColumn column in this.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
        }

        private bool isAllColumnReadOnly()
        {
            for (int i = 0; i < this.Columns.Count; i++)
            {
                if (this.Columns[i].ReadOnly == false)
                    return false;
            }
            return true;
        }

        protected override bool ProcessCmdKey(ref System.Windows.Forms.Message msg, Keys keyData)
        
        {
            // Choose a direction based on the key the user entered
            Keys direction = Keys.None;
            Keys reverseKey = Keys.None;
            switch (keyData)
            {
                case Keys.Left:              // Handle Left inside the user's text (if any)
                    if (this.EditingControl != null)
                        if (this.EditingControl.GetType() == typeof(DataGridViewTextBoxEditingControl))

                            if (((TextBox)this.EditingControl).SelectionStart > 0)
                                return base.ProcessCmdKey(ref msg, keyData);
                    direction = Keys.Shift | Keys.Tab;
                    reverseKey = Keys.Tab;
                    break;
                case Keys.Down:
                    direction = keyData;
                    reverseKey = Keys.Up;
                    break;
                case Keys.Up:
                    direction = keyData;
                    reverseKey = Keys.Down;
                    break;
                case Keys.Right:
                    // Handle Right inside the user's text (if any)
                    if (this.EditingControl != null)
                        if (this.EditingControl.GetType() == typeof(DataGridViewTextBoxEditingControl))
                            if (((TextBox)this.EditingControl).SelectionStart < this.EditingControl.Text.Length)
                                return base.ProcessCmdKey(ref msg, keyData);
                    direction = Keys.Tab;
                    reverseKey = Keys.Shift | Keys.Tab;
                    break;
                case Keys.Tab:
                    direction = keyData;
                    reverseKey = Keys.Shift | Keys.Tab;
                    break;
                case Keys.Shift | Keys.Tab:
                    direction = keyData;
                    reverseKey = Keys.Tab;
                    break;
                case Keys.Enter:
                    direction = Keys.Tab;
                    reverseKey = Keys.Shift | Keys.Tab;
                    break;
                default:
                    // Simply process all other keys normally
                    return base.ProcessCmdKey(ref msg, keyData);
            }

            if (isAllColumnReadOnly())
            {
                MyProcessCmdKey(direction);
            }
            else
            {
                if (this.CurrentCell == null && this.Rows.Count > 0)
                    this.CurrentCell = this.Rows[0].Cells[0];

                if (this.CurrentCell != null)
                {
                    int currRowIndex = this.CurrentCell.RowIndex;
                    int currColIndex = this.CurrentCell.ColumnIndex;
                    bool found = false;
                    int readonlyCount = 0;
                    switch (direction)
                    {
                        case Keys.Tab:
                            while (!found)
                            {
                                if (currColIndex == this.Columns.Count - 1)
                                {
                                    if (currRowIndex == this.Rows.Count - 1)
                                    {
                                        found = true;
                                        readonlyCount = -1;
                                    }
                                    else
                                    {
                                        currColIndex = 0;
                                    }
                                }
                                else
                                    currColIndex++;

                                if (this.Columns[currColIndex].ReadOnly == true)
                                {
                                    readonlyCount++;
                                }
                                else
                                    found = true;
                            }
                            for (int i = 0; i <= readonlyCount; i++)
                            {
                                MyProcessCmdKey(direction);
                            }
                            break;
                        case Keys.Shift | Keys.Tab:
                            while (!found)
                            {
                                if (currColIndex == 0)
                                {
                                    if (currRowIndex == 0)
                                    {
                                        found = true;
                                        readonlyCount = 1;
                                    }
                                    else
                                    {
                                        currColIndex = this.ColumnCount - 1;
                                    }
                                }
                                else
                                    currColIndex--;

                                if (this.Columns[currColIndex].ReadOnly == true)
                                {
                                    readonlyCount++;
                                }
                                else
                                    found = true;
                            }
                            for (int i = 0; i <= readonlyCount; i++)
                            {
                                MyProcessCmdKey(direction);
                            }

                            break;
                        case Keys.Up:
                        case Keys.Down:
                            MyProcessCmdKey(direction);
                            break;
                    }
                }
            }

            return true;
                //if (keyData == Keys.Tab && this[this.CurrentCell.ColumnIndex + 1, this.CurrentCell.RowIndex].ReadOnly == true) //Skip Tab order the readonly column
                //{
                //    this.CurrentCell = this[this.CurrentCell.ColumnIndex + 2, this.CurrentCell.RowIndex];
                //    return true;
                //}
                //else
                //{
                //    return base.ProcessCmdKey(ref msg, keyData);
                //}

                //if (keyData == Keys.Tab && this.EditingControl != null && msg.HWnd == this.EditingControl.Handle)
                //{
                //    return true;
                //}
                //return base.ProcessCmdKey(ref msg, keyData);
           
        }

        public void frmKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
            
            }
          

        }

        private void MyProcessCmdKey(Keys keyData)
        {
           // MessageBox.Show(this.CurrentCell.EditedFormattedValue.ToString());
            if (this.CurrentCell.EditedFormattedValue.ToString().Contains("-"))
            {
                if (this.Name == "dgvPETC0302I" || this.Name == "dgvPETC0308I")
                {
                    ProcessLeftKey(keyData);
                    ProcessRightKey(keyData);
                }
            }
            else
            {
                switch (keyData)
                {
                    case Keys.Enter:
                        ProcessTabKey(keyData);
                        break;
                    case Keys.Tab:

                        ProcessDialogKey(keyData);
                        break;
                    case Keys.Tab | Keys.Shift:
                        ProcessTabKey(keyData);
                        break;
                    case Keys.Left:

                        ProcessLeftKey(keyData);
                        break;
                    case Keys.Right:
                        ProcessRightKey(keyData);
                        break;
                    case Keys.Up:
                        ProcessUpKey(keyData);
                        break;
                    case Keys.Down:
                        ProcessDownKey(keyData);
                        break;
                    case Keys.None:
                        break;
                }
            }
        } // MyProcessCmdKey

        public void DisabledColumn(string DisablecolumnName)
        {
            this.DisablecolName = DisablecolumnName;
        }

        public void EnabledColumn(string EnablecolumnName)
        {
            this.EnablecolName = EnablecolumnName;
        }

        public void Gridcolor(Color clr)
        {
            this.GridColor = clr;
        }

        //public void Headergrid(Color clr)
        //{
        //    this.Rows.gr
        //}

        protected override void OnCellFormatting(DataGridViewCellFormattingEventArgs e)
        {
            base.OnCellFormatting(e);
            bool found = false;
            if (!string.IsNullOrWhiteSpace(DisablecolName))
            {
                if (DisablecolName.Equals("*"))
                {
                    if (e.CellStyle.BackColor != Color.Yellow)
                    {
                        e.CellStyle.BackColor = Color.FromArgb(217, 217, 217);
                        
                    }
                }
                else
                {
                    string[] arr = this.DisablecolName.Split(',');
                    foreach (string str in arr)
                    {
                        if (e.ColumnIndex == this.Columns[str].Index)
                        {
                            if (e.CellStyle.BackColor != Color.Yellow)
                                e.CellStyle.BackColor = Color.FromArgb(217, 217, 217);
                        }
                    }
                }
            }
            else if (!string.IsNullOrWhiteSpace(EnablecolName))
            {
                int i = e.ColumnIndex;
                string[] arr = this.EnablecolName.Split(',');
                foreach (string str in arr)
                {
                    if (e.ColumnIndex == this.Columns[str].Index)
                    {
                        found = true;
                    }
                }

                if (!found)
                    e.CellStyle.BackColor = Color.FromArgb(217, 217, 217);
            }
        }

        private void InitializeComponent()
        {
            ((System.ComponentModel.ISupportInitialize)(this)).BeginInit();
            this.SuspendLayout();
            // 
            // CKM_DataGridView
            // 
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.CKM_DataGridView_KeyDown);
            ((System.ComponentModel.ISupportInitialize)(this)).EndInit();
            this.ResumeLayout(false);

        }

        private   void CKM_DataGridView_KeyDown(object sender, KeyEventArgs e)
        {
            
            
        }
    

    }
}

