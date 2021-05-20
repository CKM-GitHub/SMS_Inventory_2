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
using Sms_Prog;
using SMS.CustomControls;

namespace SMS.PopUp
{
    public partial class frmPCMN0101K : SMS_Base
    {
        Login_Info loginInfo;
        int flag;
        bool allowMulticheck = false;
        int SearchMode = 1; //1 - supplier, 2 - brand,3 -- sports,4 - category,5 - pcmn0101k
        ItemSearch_Entity ise;
        PCMN0101K_BL pcmn0101kbl;
       
        public string result = string.Empty;
        public int count = 0;

        #region constructor
        /// <summary>
        /// Default
        /// </summary>
        public frmPCMN0101K()
        {
            InitializeComponent();
           
        }

        /// <summary>
        /// Get Login Info
        /// </summary>
        /// <param name="loginInfo"></param>
        /// <param name="flag"></param>
        /// <param name="allowMulticheck"></param>
        public frmPCMN0101K(Login_Info loginInfo, int flag,bool allowMulticheck)
            : base(loginInfo)
        {
            InitializeComponent();
            
            this.flag = flag;
            this.allowMulticheck = allowMulticheck;
        }

        #endregion constructor

        /// <summary>
        /// form load
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmPCMN0101K_Load(object sender, EventArgs e)
        {
            FunctionButtonDisabled(2);
            FunctionButtonDisabled(3);
            FunctionButtonDisabled(4);
            FunctionButtonDisabled(5);
            FunctionButtonDisabled(6);
            FunctionButtonDisabled(7);
            FunctionButtonDisabled(8);
            FunctionButtonDisabled(10);

            FormName = "商品検索";
            lblMode.Visible = false;
            pcmn0101kbl = new PCMN0101K_BL();
            switch (flag)
            {
                case 1:
                    txtItemcd.Focus();
                    break;
                case 2:
                    txtmakercd.Focus();
                    break;
                case 3:
                    txtJancd.Focus();
                    break;
                case 4:
                    txtskucd.Focus();
                    break;
            }
        }

        /// <summary>
        /// handle f1 to f12 click event
        /// </summary>
        /// <param name="btnIndex"></param>
        public override void functionClick(int btnIndex)
        {
            switch (btnIndex)
            {
                case 1:
                    F1();
                    break;
                case 11:
                    F11();
                    break;
                case 9:
                    F9();
                    break;
                case 12:
                    F12();
                    break;
            }
        }

        /// <summary>
        /// Bind checked data to Main Form
        /// </summary>
        private void F12()
        {
            if (gv_pcmnk.Rows.Count > 0)
            {
                bool isSelected = false;
                foreach (DataGridViewRow item in gv_pcmnk.Rows)
                {
                    isSelected = Convert.ToBoolean(item.Cells["colChk"].Value);

                    if (isSelected)
                    {
                        if (flag == 1)
                        {
                            result = item.Cells["colITEM"].Value.ToString();
                            break;
                        }
                        if (flag == 3)
                        {
                            result = item.Cells["colJANCD"].Value.ToString();
                            break;
                        }
                        if (flag == 4)
                        {
                            result = item.Cells["colSKUCD"].Value.ToString();
                            break;
                        }
                        if (flag == 2)
                        {
                            result = item.Cells["colMaker"].Value.ToString();
                            break;
                        }
                        if (flag == 5)
                        {
                            if (result != string.Empty)
                            {
                                result += ",";
                            }
                            result += item.Cells["colJANCD"].Value.ToString();
                        }
                    }
                }

                if (string.IsNullOrWhiteSpace(result))
                {
                    MessageBox.Show("Plz! Check the Checkbox.", "Message");
                }
                else
                    this.Close();
             }
            
        }

        /// <summary>
        /// Close Form
        /// </summary>
        private void F1()
        {
            this.Close();
        }

        #region Searching
        /// <summary>
        /// Display Data for GridView
        /// </summary>
        private void F11()
        {
               count = 0; // reset count
               if (!string.IsNullOrWhiteSpace(txtmakercd.Text) || !string.IsNullOrWhiteSpace(txtItemcd.Text) || !string.IsNullOrWhiteSpace(txtItemName.Text) || !string.IsNullOrWhiteSpace(txtJancd.Text) || !string.IsNullOrWhiteSpace(txtskucd.Text) ||
                   !string.IsNullOrWhiteSpace(ucBrand.UC_Code) || !string.IsNullOrWhiteSpace(ucSupplier.UC_Code) || !string.IsNullOrWhiteSpace(ucKyouGiName.UC_Code) || !string.IsNullOrWhiteSpace(ucCategory.UC_Code))
               {

                   Cursor.Current = Cursors.WaitCursor;
                   ise = GetItemSearch_Entity();
                   DataTable dt = pcmn0101kbl.PCMN0101K_Search(ise);
                   gv_pcmnk.DataSource = dt;

                   Cursor.Current = Cursors.Default;

               }
               else
               {
                   DSP_MSG("E111", "名前", string.Empty, string.Empty, string.Empty, string.Empty);
               }
        }

        /// <summary>
         /// searching
         /// </summary>
         /// <param name="sender"></param>
         /// <param name="e"></param>
        private void btnDisplay_F11_Click(object sender, EventArgs e)
        {
            F11();
        }

        #endregion

        /// <summary>
        ///  Get Search Data
        /// </summary>
        /// <returns></returns>
        private ItemSearch_Entity GetItemSearch_Entity()
        {
            ise = new ItemSearch_Entity();
            ise.makercd = txtmakercd.Text;
            ise.itemName = txtItemName.Text;
            ise.itemCode = txtItemcd.Text;
            ise.skucd = txtskucd.Text;
            ise.jancd = txtJancd.Text;
            ise.supplier = ucSupplier.UC_Code;
            ise.brand = ucBrand.UC_Code;
            ise.sports = ucKyouGiName.UC_Code;
            ise.category = ucCategory.UC_Code;
            if (chkSearch.Checked && (!string.IsNullOrWhiteSpace(txtJancd.Text) || !string.IsNullOrWhiteSpace(txtskucd.Text) || !string.IsNullOrWhiteSpace(txtItemcd.Text)))
            {
                if (rdoITEM.Checked)
                    ise.itemCheck ="1";
                else
                    ise.makerCheck ="1";
            
            }

            return ise;
        }

        /// <summary>
        /// checkbox click of gridview cell
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void gv_pcmnk_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            var senderGrid = (DataGridView)sender;
            if (e.RowIndex >= 0)
            {
                if ((sender as DataGridView).CurrentCell is DataGridViewCheckBoxCell)
                {
                        if (flag != 5)
                        {
                            foreach (DataGridViewRow row in gv_pcmnk.Rows)
                            {
                                DataGridViewCheckBoxCell chk = (DataGridViewCheckBoxCell)row.Cells["colChk"];
                                chk.Value = chk.FalseValue;
                                chk.Selected = false;
                            }
                            gv_pcmnk.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = true;
                        }
                        else
                        {
                            DataGridViewRow row = gv_pcmnk.Rows[e.RowIndex];
                             DataGridViewCheckBoxCell chk = (DataGridViewCheckBoxCell)row.Cells["colChk"];

                            if (chk.Value == null || chk.Value == chk.FalseValue)//uncheck condition
                            {
                                if (count == 11)//set maximum limit of checkbox's check
                                {
                                    chk.Value = chk.FalseValue;
                                    gv_pcmnk.RefreshEdit();//cancel checkbox's check
                                }
                                else
                                {
                                    chk.Value = chk.TrueValue;//set truevalue to true,falsevalue to false in designer
                                    count++;
                                }
                            }
                            else//check condition
                            {
                                chk.Value = chk.FalseValue;
                                count--;
                            }

                        }
                }
                    
            }
        }

        /// <summary>
        /// double click event of gridview cell
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void gv_pcmnk_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            gv_pcmnk.RefreshEdit();
        }

        #region  Handle for F9 Function
        /// <summary>
        /// handle all usercontrol to show search popup
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void uc_Enter(object sender, EventArgs e)
        {
            UC_Search ucsearch = sender as UC_Search;
            if (ucsearch.UC_Type == UC_Search.Type.Maker)
                SearchMode = 1;
            else if (ucsearch.UC_Type == UC_Search.Type.Brand)
                SearchMode = 2;
            else if (ucsearch.UC_Type == UC_Search.Type.Sports)
                SearchMode = 3;
            else if(ucsearch.UC_Type==UC_Search.Type.Category)
                SearchMode = 4;

            FunctionButtonEnabled(9);
        }

        /// <summary>
        /// reset searchmode
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void uc_Leave(object sender, EventArgs e)
        {
            if (this.ActiveControl.Name != "btnF9")
            {
                FunctionButtonDisabled(9);
                SearchMode = 5;
            }
        }

        #endregion

        #region keydown event
        private void frmPCMN0101K_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F11)
            {
                F11();
            }
        }

        #endregion

        /// <summary>
        /// show popup by SearchMode
        /// </summary>
        private void F9()
        {
            if (SearchMode == 1)
                ucSupplier.ShowSearchForm();
            else if (SearchMode == 2)
                ucBrand.ShowSearchForm();
            else if (SearchMode == 3)
                ucKyouGiName.ShowSearchForm();
            else if (SearchMode == 4)
                ucCategory.ShowSearchForm();
        }

       

    }
}
