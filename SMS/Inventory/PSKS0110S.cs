using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SMS_Entity;
using SMS.PopUp;
using SMS_BL;
using System.Threading;
using Sms_Prog;
using System.IO;
using ClosedXML.Excel;
using ElencySolutions.CsvHelper;
using System.Diagnostics;
using SMS.CustomControls;

namespace SMS.Inventory
{
    public partial class frmPSKS0110S : SMS_Base
    {
        DataTable dt_PSKS0110S = new DataTable();
        int SearchMode = 1;//1 - Suppler, 2 - Brand, 3 - Kyogi, 4 - Category
        Login_Info loginInfo;
        T_Zaiko_Entity tze;
        bool flg_Jisha_Err, flg_Toyonaka_Err = false;
        ItemSearch_Entity ise;
        PSKS0110S_BL psks0110sbl;

        #region constructor
        /// <summary>
        /// Default
        /// </summary>
        public frmPSKS0110S()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Get Login Info
        /// </summary>
        /// <param name="loginInfo"></param>
        public frmPSKS0110S(Login_Info loginInfo)
            : base(loginInfo)
        {
            InitializeComponent();
            this.loginInfo = loginInfo;
        }
        #endregion

        /// <summary>
        /// Form Load
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmPSKS0110S_Load(object sender, EventArgs e)
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

            F12_Text = "出力(F12)";
            chkMakerLikeSearch.Checked = true;
        }

        /// <summary>
        /// handle f1 to f12 click event
        /// implement base virtual function
        /// </summary>
        /// <param name="btnIndex"></param>
        public override void functionClick(int btnIndex)
        {
            switch (btnIndex)
            {
                case 1:
                    break;
                case 6:
                    Clear();
                   // dgvPSKS0110S.Enabled = false;
                    ucSupplier.SetFocus();
                    break;
                case 10:
                    if (dgvPSKS0110S.CurrentCell != null)
                        frmZaikoHenkou_Display(dgvPSKS0110S.CurrentCell.RowIndex);
                    break;
                case 9:
                    F9();
                    break;
                case 11:
                    DisplayData();
                    break;
                case 12:
                    Export();
                    break;
            }
        }

        /// <summary>
        /// F12(Export)
        /// </summary>
        private void Export()
        {
            DataTable dt = dgvPSKS0110S.DataSource as DataTable;

            if (dt != null)
            {
                DataTable dtExport = dt.Copy();

                if (dtExport.Rows.Count > 0)
                {
                    dtExport.Columns.Remove("No");
                    dtExport.Columns.Remove("VC_SHOHIN");
                    dtExport.Columns.Remove("MRenkeiDateTime");
                    dtExport.Columns.Remove("TRenkeiDateTime");
                    dtExport.Columns.Remove("ToyonakaDateTime");
                    dtExport.Columns.Remove("MakerJan");
                    dtExport.Columns.Remove("ZaikoJan");

                    DataTable dtFolder = psks0110sbl.GetFolder();
                    string Folderpath = dtFolder.Rows[0]["Char2"].ToString();
                    if (!string.IsNullOrWhiteSpace(Folderpath))
                    {
                        if (!Directory.Exists(Folderpath))
                        {
                            Directory.CreateDirectory(Folderpath);
                        }

                        #region CSV,Excel create and save
                        SaveFileDialog savedialog = new SaveFileDialog();
                        savedialog.Filter = "Csv|*.csv|Excel|*.xls";

                        savedialog.Title = "Save";
                        savedialog.FileName = dtFolder.Rows[0]["Char3"].ToString();
                        savedialog.InitialDirectory = dtFolder.Rows[0]["Char2"].ToString();
                        savedialog.RestoreDirectory = true;
                        if (savedialog.ShowDialog() == DialogResult.OK)
                        {
                            if (Path.GetExtension(savedialog.FileName).Contains("csv"))
                            {
                                CsvWriter csvwriter = new CsvWriter();
                                csvwriter.WriteCsv(dtExport, savedialog.FileName, Encoding.GetEncoding(932));
                            }
                            else
                            {
                                XLWorkbook wb = new XLWorkbook();
                                wb.Worksheets.Add(dtExport, "Sheet1");
                                wb.SaveAs(savedialog.FileName);
                            }

                            Process.Start(Path.GetDirectoryName(savedialog.FileName));
                        }
                        #endregion
                    }
                }
            }
        }

        /// <summary>
        /// Reset Data
        /// </summary>
        private void Clear()
        {
            ucSupplier.Clear();
            ucBrand.Clear();
            ucKyouGiName.Clear();
            ucCategory.Clear();

            txtItem.Text = string.Empty;
            txtMakerShohinCD.Text = string.Empty;
            txtJANCD.Text = string.Empty;
            txtShijisho.Text = string.Empty;

            cboNendo.SelectedValue = -1;
            cboSeason.SelectedValue = -1;
            cboSpecialFlag.SelectedValue = -1;
            cboYoyakuFlg.SelectedValue = -1;
          
            dgvPSKS0110S.DataSource = null;

            chkSearch.Checked = false;
            rdoItem.Checked = true;
            rdoMakerItemCD.Checked = false;
            chkWebTenpo.Checked = true;
            chkZaikoExists.Checked = true;
        }

        /// <summary>
        /// Call frmZaikoHenkou Form
        /// </summary>
        /// <param name="rowIndex"></param>
        private void frmZaikoHenkou_Display(int rowIndex)
        {
            int current_row = rowIndex;
            tze = new T_Zaiko_Entity();
            tze.HanbaiShohinCD = dt_PSKS0110S.Rows[rowIndex]["SKUCD"].ToString();
            tze.JanCD = dt_PSKS0110S.Rows[rowIndex]["JANCD"].ToString();
            tze.ItemName = dt_PSKS0110S.Rows[rowIndex]["商品名"].ToString();
            tze.Color = dt_PSKS0110S.Rows[rowIndex]["カラー"].ToString();
            tze.Size = dt_PSKS0110S.Rows[rowIndex]["サイズ"].ToString();
            tze.CapitalZaiko = dt_PSKS0110S.Rows[rowIndex]["自社"].ToString();
            tze.ToyonakaZaiko = dt_PSKS0110S.Rows[rowIndex]["豊中"].ToString();
            tze.IshibashiZaiko = dt_PSKS0110S.Rows[rowIndex]["石橋"].ToString();
            tze.EsakaZaiko = dt_PSKS0110S.Rows[rowIndex]["江坂"].ToString();
            tze.SannomiyaZaiko = dt_PSKS0110S.Rows[rowIndex]["三宮"].ToString();
            tze.HenkanNouki = dt_PSKS0110S.Rows[rowIndex]["納期"].ToString();
            frmZaikoHenkou frmZaiko = new frmZaikoHenkou(tze, loginInfo);
            frmZaiko.ShowDialog();

            if (frmZaiko.flgExist)
            {
                ucBrand.SetFocus();
            }
            else
            {
                string str = dgvPSKS0110S.Rows[current_row].Cells["colJiSha"].Value.ToString();
                if ((frmZaiko.oldJisha != "") && (dgvPSKS0110S.Rows[current_row].Cells["colJiSha"].Value.ToString() != frmZaiko.oldJisha))
                {
                    flg_Jisha_Err = true;
                }
                else
                {
                    flg_Jisha_Err = false;
                }
                if ((frmZaiko.oldToyonaka != "") && (dgvPSKS0110S.Rows[current_row].Cells["colToyonakaZaiko"].Value.ToString() != frmZaiko.oldToyonaka))
                {
                    flg_Toyonaka_Err = true;
                }
                else
                {
                    flg_Toyonaka_Err = false;
                }

                #region set zaiko
                tze = frmZaiko.tze;
                if (tze != null)
                {
                    dt_PSKS0110S.Rows[rowIndex]["自社"] = tze.CapitalZaiko;
                    dt_PSKS0110S.Rows[rowIndex]["豊中"] = tze.ToyonakaZaiko;
                    dt_PSKS0110S.Rows[rowIndex]["石橋"] = tze.IshibashiZaiko;
                    dt_PSKS0110S.Rows[rowIndex]["江坂"] = tze.EsakaZaiko;
                    dt_PSKS0110S.Rows[rowIndex]["三宮"] = tze.SannomiyaZaiko;
                    dt_PSKS0110S.Rows[rowIndex]["納期"] = tze.HenkanNouki;
                }
                #endregion

                if (flg_Jisha_Err)
                    dgvPSKS0110S.Rows[current_row].Cells["colJiSha"].Style.BackColor = Color.Yellow;
                if (flg_Toyonaka_Err)
                    dgvPSKS0110S.Rows[current_row].Cells["colToyonakaZaiko"].Style.BackColor = Color.Yellow;
            }
        }

        /// <summary>
        /// Searching
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDisplay_Click(object sender, EventArgs e)
        {
            DisplayData();
        }

        /// <summary>
        /// Display Data for GridView
        /// repaied by PTK 02182019
        /// </summary>
        private void DisplayData()
        {
            if (ErrorCheck())
            {
                Cursor.Current = Cursors.WaitCursor;
                tze = new T_Zaiko_Entity();
                string qty = tze.CapitalZaiko;
                ise = GetData();
                dt_PSKS0110S = psks0110sbl.PSKS0110S_Select(ise);

                if (dt_PSKS0110S.Rows.Count > 0)
                {
                    //if(dt_PSKS0110S !=null || dt_PSKS0110S.Rows.Count>0)
                    //{
                    dgvPSKS0110S.DataSource = dt_PSKS0110S;
                    Cursor.Current = Cursors.Default;
                    btnDisplay.Focus();
                }
                else
                {
                    DSP_MSG("E128", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                }

            }
        }

        /// <summary>
        /// Get Search Data
        /// </summary>
        /// <returns></returns>
        public ItemSearch_Entity GetData()
        {
            ise = new ItemSearch_Entity();

            ise.itemName = txtItemName.Text;

            try
            {
                ise.specialFlag = cboSpecialFlag.SelectedValue.ToString();
               
            }catch
            {
                ise.specialFlag = "-1";
            }
            try{ise.yoyakuFlg = cboYoyakuFlg.SelectedValue.ToString(); }
            catch
            {
                ise.yoyakuFlg = "-1";
            }
            try{
                ise.nendo = cboNendo.SelectedValue.ToString();
            }
            catch
            {
                ise.nendo = "-1";
            }try{
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
            //ise.nendo = cboNendo.SelectedValue.ToString();
            //ise.season = cboSeason.SelectedValue.ToString();
            ise.catalog = txtCatalog.Text;
            ise.shijisho = txtShijisho.Text;
            ise.makercd = txtMakerShohinCD.Text.TrimStart(',').TrimEnd(',').Trim().Replace("\r\n", string.Empty);
            ise.itemCode = txtItem.Text.TrimStart(',').TrimEnd(',').Trim().Replace("\r\n", string.Empty);
            ise.jancd = txtJANCD.Text.TrimStart(',').TrimEnd(',').Trim().Replace("\r\n", string.Empty); ;
            ise.webTenpo = chkWebTenpo.Checked ? "1" : "0";
            ise.zaikoExists = chkZaikoExists.Checked ? "1" : "0";
            ise.skucd = ucSKUCD.UC_Code;

            if (chkSearch.Checked && (!string.IsNullOrWhiteSpace(txtItem.Text) || !string.IsNullOrWhiteSpace(ucSKUCD.UC_Code) || !string.IsNullOrWhiteSpace(txtJANCD.Text)))
            {
                ise.itemCheck = rdoItem.Checked ? "1" : string.Empty;
                ise.makerCheck = rdoMakerItemCD.Checked ? "1" : string.Empty;
            }

            if (!string.IsNullOrWhiteSpace(txtMakerShohinCD.Text))
                ise.makerLikeSearch = chkMakerLikeSearch.Checked ? "1" : string.Empty;

            return ise;
        }

        /// <summary>
        /// handle for Searching with function key
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmPSKS0110S_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F11)
            {

                    DisplayData();
            }
        }

        /// <summary>
        /// Error Check function
        /// </summary>
        /// <returns></returns>
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
                string.IsNullOrWhiteSpace(txtJANCD.Text) && 
                string.IsNullOrWhiteSpace(ucSKUCD.UC_Code) )            
            {
                DSP_MSG("E111", "名前", string.Empty, string.Empty, string.Empty, string.Empty);
                return false;
            }
            else return true;
            #endregion
        }

        /// <summary>
        /// Change Back Cell Color to yellow 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvPSKS0110S_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if ((string.IsNullOrWhiteSpace(dgvPSKS0110S.Rows[e.RowIndex].Cells["colTRenkei"].Value.ToString())) && (!string.IsNullOrWhiteSpace(dgvPSKS0110S.Rows[e.RowIndex].Cells["colZaikoJan"].Value.ToString())))
            {
                dgvPSKS0110S.Rows[e.RowIndex].Cells["colJiSha"].Style.BackColor = Color.Yellow;
            }
            if ((string.IsNullOrWhiteSpace(dgvPSKS0110S.Rows[e.RowIndex].Cells["colMRenkei"].Value.ToString())) && (!string.IsNullOrWhiteSpace(dgvPSKS0110S.Rows[e.RowIndex].Cells["colMakerJan"].Value.ToString())))
            {
                dgvPSKS0110S.Rows[e.RowIndex].Cells["colMaker"].Style.BackColor = Color.Yellow;
            }
            if ((string.IsNullOrWhiteSpace(dgvPSKS0110S.Rows[e.RowIndex].Cells["colToyonaka"].Value.ToString())) && (!string.IsNullOrWhiteSpace(dgvPSKS0110S.Rows[e.RowIndex].Cells["colZaikoJan"].Value.ToString())))
            {
                dgvPSKS0110S.Rows[e.RowIndex].Cells["colToyonakaZaiko"].Style.BackColor = Color.Yellow;
            }
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
