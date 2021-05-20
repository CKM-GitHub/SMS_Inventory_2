using SMS.CustomControls;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using System.Collections;

namespace SMS
{
    public class SMS_GridView : DataGridView
    {
        public bool IsCellValid { get; set; }
        string DisablecolName = string.Empty;
        string EnablecolName = string.Empty;

        public SMS_GridView()
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


            this.Visible = true;
            
        }

        protected void HeaderColumnRowVisible(bool visible)
        {
            if (visible)
            {
                this.ColumnHeadersVisible = true;
                this.RowHeadersVisible = true;
            }
            else
            {
                this.ColumnHeadersVisible = false;
                this.RowHeadersVisible = false;
            }
            this.ScrollBars = ScrollBars.None; 
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
                if (this.CurrentCell == null && this.Rows.Count>0)
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


            //if (this.CurrentCell == null)
            //    this.CurrentCell = this.Rows[0].Cells[0];

            //// Setup to move to the next non-ReadOnly cell if the entered data is valid
            //int currRowIndex = this.CurrentCell.RowIndex;
            //int currColIndex = this.CurrentCell.ColumnIndex;
            //IsCellValid = true;

            //// Process the command key to trigger CellValidating and move to the next cell (which will be in whatever 
            //// direction the command key dictates).  If validation failed, simply return to stay in the original cell.
            //MyProcessCmdKey(direction);
            //if (IsCellValid == false) return true;

            //// If the (new) current cell is ReadOnly and there are no more non-ReadOnly cells in the direction of the command
            //// key (signalled by 0 cellsToSkip), we need to go back to the original cell.
            //int cellsToSkip = GetSkipReadOnlyCellCount(direction, currRowIndex, currColIndex);
            //if (this.CurrentCell.ReadOnly == true && cellsToSkip == 0) MyProcessCmdKey(reverseKey);

            //// Skip the ReadOnly cells, if any, in the direction of the command key (provided there's a remaining non-ReadOnly
            //// cell in the direction of the command key).
            //for (int i = 0; i < cellsToSkip; i++) MyProcessCmdKey(direction);

            //// Indicate that we've handled the command key
            return true;
        } // ProcessCmdKey

        private bool isAllColumnReadOnly()
        {
            for (int i = 0; i < this.Columns.Count; i++)
            {
                if (this.Columns[i].ReadOnly == false)
                    return false;
            }
            return true;
        }

        private void MyProcessCmdKey(Keys keyData)
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
        } // MyProcessCmdKey

        /// <summary>
        /// Counts the number of ReadOnly dgv cells that must be skipped based on the starting cell and the direction 
        /// implied by the command key used to exit the current cell (if any).
        /// </summary>
        public int GetSkipReadOnlyCellCount(Keys direction, int currRowIndex, int currColIndex)
        {
            // The next cell direction is set by the caller based on the command key the user entered to signal data 
            // entry, if any.  If the user clicked off the cell instead of using a command key, the direction will be 
            // None.  All command keys/clicks are translated to one of these 5 values.

            int skipReadOnlyCellCount = 0;
            int startingRowIndex;
            int startingColIndex;
            bool nonReadOnlyCellFound = false;
            switch (direction)
            {
                case Keys.Shift | Keys.Tab:

                    //case Keys.Left

                    // Find the next non-ReadOnly cell moving Left and Up
                    if (currColIndex > 0) startingColIndex = currColIndex - 1;
                    else
                    {
                        startingColIndex = this.ColumnCount;
                        if (currRowIndex > 0) startingRowIndex = currRowIndex - 1;
                    }
                    for (int tagRowIndex = currRowIndex; tagRowIndex >= 0; tagRowIndex--)
                    {
                        for (int tagColIndex = startingColIndex; tagColIndex >= 0; tagColIndex--)
                        {
                            if (this[tagColIndex, tagRowIndex].ReadOnly == false)
                            {
                                nonReadOnlyCellFound = true;
                                break;
                            }
                            skipReadOnlyCellCount++;
                        }
                        if (nonReadOnlyCellFound == true) break;
                        startingColIndex = this.Columns.Count - 1;
                    }
                    // There are no more non-ReadOnly cells, so don't count any of the RaedOnly ones
                    if (nonReadOnlyCellFound == false) skipReadOnlyCellCount = 0;
                    break;

                case Keys.Tab:
                    //case Keys.Right:
                    // Find the next non-ReadOnly cell moving Right and Down
                    if (currColIndex < this.ColumnCount - 1) startingColIndex = currColIndex + 1;
                    else
                    {
                        startingColIndex = 0;
                        if (currRowIndex < this.RowCount - 1) startingRowIndex = currRowIndex + 1;
                    }
                    for (int tagRowIndex = currRowIndex; tagRowIndex < this.Rows.Count; tagRowIndex++)
                    {
                        for (int tagColIndex = startingColIndex; tagColIndex < this.Columns.Count; tagColIndex++)
                        {
                            if (this[tagColIndex, tagRowIndex].ReadOnly == false)
                            {
                                nonReadOnlyCellFound = true;
                                break;
                            }
                            skipReadOnlyCellCount++;
                        }
                        if (nonReadOnlyCellFound == true) break;
                        startingColIndex = 0;
                    }
                    // There are no more non-ReadOnly cells, so don't count any of the RaedOnly ones
                    if (nonReadOnlyCellFound == false) skipReadOnlyCellCount = 0;
                    break;

                case Keys.Up:
                    // Find the next non-ReadOnly cell moving straight Up
                    startingRowIndex = Math.Max(currRowIndex - 1, 0);
                    for (int tagRowIndex = startingRowIndex; tagRowIndex >= 0; tagRowIndex--)
                    {
                        if (this[currColIndex, tagRowIndex].ReadOnly == false)
                        {
                            nonReadOnlyCellFound = true;
                            break;
                        }
                        skipReadOnlyCellCount++;
                    }
                    // There are no more non-ReadOnly cells, so don't count any of the RaedOnly ones
                    if (nonReadOnlyCellFound == false) skipReadOnlyCellCount = 0;
                    break;

                case Keys.Down:
                    // Find the next non-ReadOnly cell moving straight Down
                    startingRowIndex = Math.Min(currRowIndex + 1, this.Rows.Count - 1);
                    for (int tagRowIndex = startingRowIndex; tagRowIndex < this.Rows.Count; tagRowIndex++)
                    {
                        if (this[currColIndex, tagRowIndex].ReadOnly == false)
                        {
                            nonReadOnlyCellFound = true;
                            break;
                        }
                        skipReadOnlyCellCount++;
                    }
                    // There are no more non-ReadOnly cells, so don't count any of the RaedOnly ones
                    if (nonReadOnlyCellFound == false) skipReadOnlyCellCount = 0;
                    break;

                case Keys.None:
                    skipReadOnlyCellCount = 0;
                    break;
            }

            return skipReadOnlyCellCount;
        }  // GetSkipReadOnlyCellCount

        int Key_Check = 0;
        bool initial_Check = true;
        DataGridViewMaskedTextControl tc;

        protected override void OnEditingControlShowing(DataGridViewEditingControlShowingEventArgs e)
        {
            base.OnEditingControlShowing(e);
            if (e.Control.GetType() == typeof(DataGridViewMaskedTextControl))
                tc = e.Control as DataGridViewMaskedTextControl;
            else
                tc = null;
        }

        //public void SetReadOnly(DataGridViewColumn col,bool value)
        //{
        //    col.ReadOnly = value;
        //    if (value)
        //        col.DefaultCellStyle.BackColor = Color.FromArgb(255, 255, 153);
        //    else
        //        col.DefaultCellStyle.BackColor = Color.White;          
        //}

        //protected override bool ProcessDialogKey(Keys keyData)
        //{
        //    Keys key = (keyData & Keys.KeyCode);
        //    if (key == Keys.Enter)
        //    {
        //        base.ProcessTabKey(Keys.Tab);//make enter key like tab key
        //        return true;
        //    }
        //    else if (key == Keys.Escape)
        //    {
        //        tc = null;//remove validate on escape key press
        //    }
        //    return base.ProcessDialogKey(keyData);
        //}

        //protected override bool ProcessDataGridViewKey(KeyEventArgs e)
        //{
        //    if (e.KeyCode == Keys.Enter)
        //    {
        //        base.ProcessTabKey(Keys.Tab);
        //        return true;
        //    }
        //    return base.ProcessDataGridViewKey(e);
        //}

        //protected override void OnCellEndEdit(DataGridViewCellEventArgs e)
        //{
        //    base.OnCellEndEdit(e);
        //    if(this.DataSource != null)
        //        this.BindingContext[this.DataSource].EndCurrentEdit();
        //}

        protected override void OnCellValidating(DataGridViewCellValidatingEventArgs e)
        {
            base.OnCellValidating(e);
            if (tc != null && !string.IsNullOrWhiteSpace((tc.Text.Replace("/", "")).Replace(":", "")))
            {
                if (!tc.IsValid())
                {
                    e.Cancel = true;
                    tc.SelectAll();
                    Key_Check = 1;
                }
                else if (initial_Check)
                {
                    //tc = null;
                    e.Cancel = false;
                }
                else
                {
                    e.Cancel = true;
                }
            }
            //else if (!initial_Check)
            //{
            //    e.Cancel = true;
            //}
            //initial_Check = true;
        }

        //protected override void OnKeyDown(KeyEventArgs e)
        //{
        //    //if (e.KeyData == Keys.Tab)
        //    //    EnterPress();
        //    //else
        //    if (e.KeyData != Keys.Enter)//prevent navigate to new row on enter press
        //        base.OnKeyDown(e);
        //}

        ////protected override void OnCellEnter(DataGridViewCellEventArgs e)
        ////{
        ////    base.OnCellEnter(e);
        ////    if (this.Columns[e.ColumnIndex].ReadOnly)//Skip Tab Index to readonly column
        ////    {
        ////        SendKeys.Send("{Tab}");
        ////    }
        ////}

        /// <summary>
        /// multiple columns with separated by comma,if all col Disable send param as *
        /// </summary>
        /// <param name="columnName"></param>
        public void DisabledColumn(string DisablecolumnName)
        {
            this.DisablecolName = DisablecolumnName;
        }

        public void EnabledColumn(string EnablecolumnName)
        {
            this.EnablecolName = EnablecolumnName;
        }

        protected override void OnCellFormatting(DataGridViewCellFormattingEventArgs e)
        {
            base.OnCellFormatting(e);
            bool found = false;
            if (!string.IsNullOrWhiteSpace(DisablecolName))
            {
                if (DisablecolName.Equals("*"))
                {
                    if (e.CellStyle.BackColor != Color.Yellow)
                        e.CellStyle.BackColor = Color.FromArgb(217, 217, 217);
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

                if(!found && e.CellStyle.BackColor!=Color.Red)
                    e.CellStyle.BackColor = Color.FromArgb(217, 217, 217);
            }
        }

        public void DisabledSorting()
        {
           foreach (DataGridViewColumn column in this.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
        }
       
    }
}


