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
using System.IO;
using CsvHelper;
using ClosedXML.Excel;
using ElencySolutions.CsvHelper;
using System.Diagnostics;
using SMS.CustomControls;
using System.Collections;

namespace SMS.PopUp
{
    public partial class FrmPSKS0110S_Search : SMS_Base
    {
        DataTable dt_PSKS0110S = new DataTable();
        int SearchMode = 1;//1 - Suppler, 2 - Brand, 3 - Kyogi, 4 - Category
        Login_Info loginInfo;
        T_Zaiko_Entity tze;
        //bool flg_Jisha_Err, flg_Toyonaka_Err = false;
        ItemSearch_Entity ise;
        PSKS0110S_BL psks0110sbl;

        public string result = string.Empty;
        ArrayList ArradminCD = new ArrayList();


        public FrmPSKS0110S_Search()
        {
            InitializeComponent();
        }

        private void FrmPSKS0110S_Search_Load(object sender, EventArgs e)
        {
            psks0110sbl = new PSKS0110S_BL();

            FormName = "在庫照会";
            lblMode.Visible = false;

            FunctionButtonDisabled(2);
            FunctionButtonDisabled(3);
            FunctionButtonDisabled(5);
            FunctionButtonDisabled(4);
            FunctionButtonDisabled(7);
            FunctionButtonDisabled(8);
            FunctionButtonDisabled(9);
            FunctionButtonDisabled(11);

            dgvPSKS0110S.DisabledColumn("*");
            dgvPSKS0110S.DisabledSorting();

            cboNendo.Bind();
            cboSeason.Bind();
            cboSpecialFlag.Bind();
            cboYoyakuFlg.Bind();

            F12_Text = "対象に追加(F12)";
            chkMakerLikeSearch.Checked = true;
            F10_Text = "表示(F10)";
            F1_Text = "戻る(F1)";
            F9_Text = "検索(F9)";
            //panelTitleRight.Enabled = false;
            lblSOperator.Visible = false;
            lblSProcessingDate.Visible = false;
            smS_Label1.Visible = false;         
        }

        public override void functionClick(int btnIndex)
        {
            switch (btnIndex)
            {
                case 1:
                    break;
                case 6:
                    Clear();
                    ucSupplier.SetFocus();
                    break;
                case 10:
                    DisplayData();
                    break;
                case 9:
                    F9();
                    break;
                case 11:            
                    break;
                case 12:
                    F12();
                    break;
            }
        }

        private void F12()
        {
             if ( dgvPSKS0110S.Rows.Count > 0)
             {
                //bool isSelected = false;
                //foreach (DataGridViewRow item in dgvPSKS0110S.Rows)
                //{
                //    isSelected = Convert.ToBoolean(item.Cells["colChk"].Value);
                //    if (isSelected)
                //    {                       
                //       if (result != string.Empty)                    
                //       {
                //           result += ",";
                //       }
                //       result += item.Cells["colAdmin"].Value.ToString();                    
                //    }
                //}
                 if(ArradminCD.Count>0)
                     result = string.Join(",", (string[])ArradminCD.ToArray(Type.GetType("System.String")));

                if (string.IsNullOrWhiteSpace(result))
                {
                    //MessageBox.Show("Plz! Check the Checkbox.", "Message");
                    DSP_MSG("E805", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                }
                else
                    this.Close();
            }
        }

        private void Clear()
        {
            ucSupplier.Clear();
            ucBrand.Clear();
            ucKyouGiName.Clear();
            ucCategory.Clear();

            txtItem.Text = string.Empty;
            txtItemName.Text = string.Empty;
            txtMakerShohinCD.Text = string.Empty;
            txtJANCD.Text = string.Empty;
            txtShijisho.Text = string.Empty;
            txtCatalog.Text = string.Empty;

            cboNendo.SelectedValue = -1;
            cboSeason.SelectedValue = -1;
            cboSpecialFlag.SelectedValue = -1;
            cboYoyakuFlg.SelectedValue = -1;

            dgvPSKS0110S.DataSource = null;

        }

        public ItemSearch_Entity GetData()
        {
            ise = new ItemSearch_Entity();

            ise.itemName = txtItemName.Text;

            try
            {
                ise.specialFlag = cboSpecialFlag.SelectedValue.ToString();
            }
            catch
            {
                ise.specialFlag = "-1";
            }
            try 
            { ise.yoyakuFlg = cboYoyakuFlg.SelectedValue.ToString(); }
            catch
            {
                ise.yoyakuFlg = "-1";
            }
            try
            {
                ise.nendo = cboNendo.SelectedValue.ToString();
            }
            catch
            {
                ise.nendo = "-1";
            } try
            {
                ise.season = cboSeason.SelectedValue.ToString();
            }
            catch
            {
                ise.season = "-1";
            }

            ise.sports = ucKyouGiName.UC_Code;
            ise.bunrui = ucCategory.UC_Code;
            ise.brand = ucBrand.UC_Code;
            ise.supplier = ucSupplier.UC_Code;
            ise.catalog = txtCatalog.Text;
            ise.shijisho = txtShijisho.Text;
            ise.makercd = txtMakerShohinCD.Text.TrimStart(',').TrimEnd(',').Trim().Replace("\r\n", string.Empty);
            ise.itemCode = txtItem.Text.TrimStart(',').TrimEnd(',').Trim().Replace("\r\n", string.Empty);
            ise.jancd = txtJANCD.Text.TrimStart(',').TrimEnd(',').Trim().Replace("\r\n", string.Empty); ;

            if (!string.IsNullOrWhiteSpace(txtMakerShohinCD.Text))
                ise.makerLikeSearch = chkMakerLikeSearch.Checked ? "1" : string.Empty;
                     
            return ise;
        }

        private void DisplayData()
        {
            if (ErrorCheck())
            {
                Cursor.Current = Cursors.WaitCursor;
                //tze = new T_Zaiko_Entity();
                //string qty = tze.CapitalZaiko;
                ise = GetData();
                dt_PSKS0110S = psks0110sbl.PSKS0110S_Search_Select(ise);
                dgvPSKS0110S.DataSource = dt_PSKS0110S;
                Cursor.Current = Cursors.Default;
                btnDisplay.Focus();
            }
        }

        private bool ErrorCheck()
        {
            #region check SearchItem exists
            if (string.IsNullOrWhiteSpace(txtItem.Text) &&
                cboSpecialFlag.SelectedValue.Equals("-1") &&
                cboYoyakuFlg.SelectedValue.Equals("-1") &&
                cboNendo.SelectedValue.Equals("-1") &&
                cboSeason.SelectedValue.Equals("-1") &&
                string.IsNullOrWhiteSpace(ucSupplier.UC_Code) &&
                string.IsNullOrWhiteSpace(ucBrand.UC_Code) &&
                string.IsNullOrWhiteSpace(ucKyouGiName.UC_Code) &&
                string.IsNullOrWhiteSpace(ucCategory.UC_Code) &&
                string.IsNullOrWhiteSpace(txtCatalog.Text) &&
                string.IsNullOrWhiteSpace(txtShijisho.Text) &&
                string.IsNullOrWhiteSpace(txtMakerShohinCD.Text) &&
                string.IsNullOrWhiteSpace(txtItem.Text) &&
                string.IsNullOrWhiteSpace(txtJANCD.Text))
               
            {
                DSP_MSG("E111", "名前", string.Empty, string.Empty, string.Empty, string.Empty);
                return false;
            }
            else return true;
            #endregion
        }

        private void btnSelectAll_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row1 in dgvPSKS0110S.Rows)
            {
                if (Convert.ToBoolean(row1.Cells["colChk"].EditedFormattedValue) == false)
                {
                    row1.Cells["colChk"].Value = true;
                    ArradminCD.Add(row1.Cells["colAdmin"].Value.ToString());
                }
            }
        }

        private void btnUnselectAll_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row1 in dgvPSKS0110S.Rows)
            {
                if (Convert.ToBoolean(row1.Cells["colChk"].EditedFormattedValue) == true)
                {
                    row1.Cells["colChk"].Value = false;
                    ArradminCD.Remove(row1.Cells["colAdmin"].Value.ToString());
                }
            }
        }

        private void btnDisplay_Click(object sender, EventArgs e)
        {
            DisplayData();
        }

        #region Handle for F9 Function
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
            else if (ucsearch.UC_Type == UC_Search.Type.Category)
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
            if (this.ActiveControl.Name != "btnF9")//if btnf9 click show popup,else reset SearchMode
            {
                FunctionButtonDisabled(9);
                SearchMode = 0;
            }
        }
        #endregion

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

        private void dgvPSKS0110S_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex > 0 && e.RowIndex >= 0)
            {
                if ((sender as DataGridView).CurrentCell is DataGridViewCheckBoxCell)
                {
                    if ((Convert.ToBoolean(dgvPSKS0110S.Rows[e.RowIndex].Cells[e.ColumnIndex].EditedFormattedValue) == true))
                    {
                        DataGridViewCheckBoxCell chk1 = dgvPSKS0110S.Rows[e.RowIndex].Cells["colChk"] as DataGridViewCheckBoxCell;
                        ArradminCD.Add(dgvPSKS0110S.Rows[e.RowIndex].Cells["colAdmin"].Value.ToString());

                    }
                    else
                    {
                        ArradminCD.Remove(dgvPSKS0110S.Rows[e.RowIndex].Cells["colAdmin"].Value.ToString());
                    }
                }

            }
        }
      
    }
}
