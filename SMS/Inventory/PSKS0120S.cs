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
    public partial class frmPSKS0120S : SMS_Base
    {
        Login_Info loginInfo;
        PSKS0120S_BL p102Sbl;
        M_Zaiko_Kanri_Entity zaikokanri;
        ItemSearch_Entity ise;
        L_Log_Entity lle;
        DataTable dtPSKS0102S;
        string OperatorCD = string.Empty;
        int SearchMode = 1;//1 - Suppler, 2 - Brand, 3 - Kyogi, 4 - Category
        DataTable dt;

        public frmPSKS0120S()
        {
            InitializeComponent();
        }

        public frmPSKS0120S(Login_Info loginInfo)
            : base(loginInfo)
        {
            InitializeComponent();
            this.loginInfo = loginInfo;
            p102Sbl = new PSKS0120S_BL();
        }

        private void frmPSKS0120S_Load(object sender, EventArgs e)
        {
            FormName = "在庫状況管理照会";
            lblMode.Visible = false;

            FunctionButtonDisabled(2);
            FunctionButtonDisabled(3);
            FunctionButtonDisabled(5);
            FunctionButtonDisabled(4);
            //FunctionButtonDisabled(6);
            FunctionButtonDisabled(7);
            FunctionButtonDisabled(8);
            FunctionButtonDisabled(9);
            FunctionButtonDisabled(10);

            dgvPSKS0102S.EnabledColumn("colShijisho,colMemo,colCapitalZaiko,colToyonakaZaiko");
            dgvPSKS0102S.DisabledSorting();

            cboNendo.Bind();
            cboSeason.Bind();
            cboSpecialFlag.Bind();
            cboYoyakuFlg.Bind();

            F11_Text = "Excel出力(F11)";
            chkMakerLikeSearch.Checked = true;
            Initial_GridviewBind();
        }


        private void Initial_GridviewBind()
        {
            dtPSKS0102S = new DataTable();
            zaikokanri = new M_Zaiko_Kanri_Entity();
            zaikokanri.OperatorCD = loginInfo.OperatorCode;
            dtPSKS0102S = p102Sbl.PSKS0120S_InitialSelect(zaikokanri);
            if (dtPSKS0102S.Rows.Count > 0)
            {
                dgvPSKS0102S.DataSource = dtPSKS0102S;
            }
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
                    ucSupplier.SetFocus();
                    break;
                case 11:
                    Export();
                    break;
                case 12:
                    UpdateData();
                    break;
            }
        }

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
            txtCatalog.Text = string.Empty;

            cboNendo.SelectedValue = -1;
            cboSeason.SelectedValue = -1;
            cboSpecialFlag.SelectedValue = -1;
            cboYoyakuFlg.SelectedValue = -1;

            DisplayData();
        }

        /// <summary>
        /// F11(Export)
        /// </summary>
        private void Export()
        {
            dt = dgvPSKS0102S.DataSource as DataTable;

            if (dt != null)
            {
                DataTable dtExport = dt.Copy();

                if (dtExport.Rows.Count > 0)
                {
                    dtExport.Columns.Remove("No");
                    dtExport.Columns.Remove("OperatorCD");
                    dtExport.Columns.Remove("AdminCD");
                    dtExport.Columns.Remove("Zaiko_keisan");
                    dtExport.Columns["vm_shiiresaki"].ColumnName = "仕入先名";
                    dtExport.Columns["vm_brand"].ColumnName = "ブランド名";
                    dtExport.Columns["vc_maker_shohin"].ColumnName = "メーカー商品CD";
                    dtExport.Columns["vm_kihon_shohin"].ColumnName = "商品名";
                    dtExport.Columns["vm_color"].ColumnName = "カラー";
                    dtExport.Columns["vm_size"].ColumnName = "サイズ";
                    dtExport.Columns["CapitalZaiko"].ColumnName = "Web在庫";
                    dtExport.Columns["ToyonakaZaiko"].ColumnName = "豊中";
                    dtExport.Columns["Kijunsu"].ColumnName = "基準数";
                    dtExport.Columns["Shijisho"].ColumnName = "指示書番号";
                    dtExport.Columns["Memo"].ColumnName = "備考";

                    DataTable dtFolder = p102Sbl.GetFolder();
                    string Folderpath = dtFolder.Rows[0]["Char2"].ToString();
                    if (!string.IsNullOrWhiteSpace(Folderpath))
                    {
                        if (!Directory.Exists(Folderpath))
                        {
                            Directory.CreateDirectory(Folderpath);
                        }

                        #region CSV,Excel create and save00
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

        private void btnDisplay_Click(object sender, EventArgs e)
        {
            DisplayData();
        }

        private void DisplayData()
        {
            Cursor.Current = Cursors.WaitCursor;
            dtPSKS0102S = new DataTable();
            OperatorCD = loginInfo.OperatorCode;
            ise = Getdata();
            dtPSKS0102S = new DataTable();
            dtPSKS0102S = p102Sbl.PSKS0120S_Select(ise, OperatorCD);
            dgvPSKS0102S.DataSource = dtPSKS0102S;
            Cursor.Current = Cursors.Default;
            btnDisplay.Focus();
        }

        public ItemSearch_Entity Getdata()
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
             try { ise.yoyakuFlg = cboYoyakuFlg.SelectedValue.ToString(); }
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
             //ise.nendo = cboNendo.SelectedValue.ToString();
             //ise.season = cboSeason.SelectedValue.ToString();
             ise.catalog = txtCatalog.Text;
             ise.shijisho = txtShijisho.Text;
             ise.makercd = txtMakerShohinCD.Text.TrimStart(',').TrimEnd(',').Trim().Replace("\r\n", string.Empty);
             ise.itemCode = txtItem.Text.TrimStart(',').TrimEnd(',').Trim().Replace("\r\n", string.Empty);
             ise.jancd = txtJANCD.Text.TrimStart(',').TrimEnd(',').Trim().Replace("\r\n", string.Empty); 

             if (!string.IsNullOrWhiteSpace(txtMakerShohinCD.Text))
                 ise.makerLikeSearch = chkMakerLikeSearch.Checked ? "1" : string.Empty;

             return ise;
        }

        private void UpdateData()
        {
            if ((DSP_MSG("Q101", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty).Equals("1")))
            {
                DataTable dtUpdate = new DataTable();
                dtUpdate.Columns.Add("OperatorCD");
                dtUpdate.Columns.Add("AdminCD");
                dtUpdate.Columns.Add("Shijisho");
                dtUpdate.Columns.Add("Memo");

                dt = dgvPSKS0102S.DataSource as DataTable;
                foreach (DataRow dr in dt.Rows)
                {
                    if (dr.RowState == DataRowState.Modified)
                    {
                        DataRow row = dtUpdate.NewRow();
                        row["OperatorCD"] = dr["OperatorCD"];
                        row["AdminCD"] = dr["AdminCD"];
                        row["Shijisho"] = dr["Shijisho"];
                        row["Memo"] = dr["Memo"];
                        dtUpdate.Rows.Add(row);
                    }
                }

                if (dtUpdate.Rows.Count > 0)
                {
                    zaikokanri = new M_Zaiko_Kanri_Entity();
                    zaikokanri.xml = p102Sbl.DataTableToXml(dtUpdate);
                    lle = GetLogData();
                    if (p102Sbl.PSKS0120S_Update(zaikokanri,lle))
                    {
                        DSP_MSG("I101", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                    }
                }
            }

        }

        private L_Log_Entity GetLogData()
        {
            lle = new L_Log_Entity();
            lle.InsertOperator = loginInfo.OperatorCode;
            lle.Program = "PSKS0102S";
            lle.PC = loginInfo.PcName;
            lle.OperateMode = loginInfo.StartMode;
            lle.KeyItem = txtJANCD.Text;

            return lle;
        }

        private void frmPSKS0120S_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F10)
            {
                DisplayData();
            }
        }

        private void dgvPSKS0102S_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (dgvPSKS0102S.Rows[e.RowIndex].Cells["colZaiko_keisan"].Value.ToString() == "1")
            {
                dgvPSKS0102S.Rows[e.RowIndex].Cells["colToyonakaZaiko"].Style.BackColor = Color.FromArgb(252, 228, 214);
            }
            else dgvPSKS0102S.Rows[e.RowIndex].Cells["colToyonakaZaiko"].Style.BackColor = Color.FromArgb(217, 217, 217);

            if((dgvPSKS0102S.Rows[e.RowIndex].Cells["colZaiko_keisan"].Value.ToString() == "0") ||
            (dgvPSKS0102S.Rows[e.RowIndex].Cells["colZaiko_keisan"].Value.ToString() == "1"))
            {
                dgvPSKS0102S.Rows[e.RowIndex].Cells["colCapitalZaiko"].Style.BackColor = Color.FromArgb(252, 228, 214);
            }
            else dgvPSKS0102S.Rows[e.RowIndex].Cells["colCapitalZaiko"].Style.BackColor = Color.FromArgb(217, 217, 217);
        }

        private void dgvPSKS0102S_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == dgvPSKS0102S.Columns["colMemo"].Index)
                dgvPSKS0102S.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
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
