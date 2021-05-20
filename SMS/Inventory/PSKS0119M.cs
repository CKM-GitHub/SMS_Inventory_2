using SMS_Entity;
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
using System.Data;
using System.Threading;
using Sms_Prog;
using SMS.CustomControls;
using System.IO;
using ExcelDataReader;
using SMS.PopUp;
using ClosedXML.Excel;
using System.Collections;

namespace SMS.Inventory
{
    public partial class frmPSKS0119M : SMS_Base
    {
        Login_Info loginInfo;
        PSKS0119M_BL psks0119mbl = new PSKS0119M_BL();
        M_Zaiko_Kanri_Entity mzke;
        L_Log_Entity lle;
        string resultAdminCD = string.Empty;
        string adminCD = string.Empty;
        ArrayList ArradminCD = new ArrayList();

        DataTable dtSearch = new DataTable();
        DataTable dtGridData = new DataTable();

        public frmPSKS0119M()
        {
            InitializeComponent();
        }

        #region constructor

        public frmPSKS0119M(Login_Info loginInfo)
            : base(loginInfo)
        {            
            InitializeComponent();
            this.loginInfo = loginInfo;
        }
        #endregion

        private void frmPSKS0119M_Load(object sender, EventArgs e)
        {
            this.FormName = "在庫状況管理登録";
            lblMode.Visible = false;
            FunctionButtonDisabled(2);
            FunctionButtonDisabled(3);
            FunctionButtonDisabled(5);
            FunctionButtonDisabled(7);
            FunctionButtonDisabled(8);
            FunctionButtonDisabled(9);
            FunctionButtonDisabled(10);
            txtFilePath.Enabled = false;

            gdvPSKS0119M.EnabledColumn("colChk,colRefNo,colInstructionNo,colRemark");
            gdvPSKS0119M.DisabledSorting();

            cboTaisho.Bind();
            cboTaisho1.Bind();
            cboShiiresaki.Bind();
            cboBrand.Bind();
            cboZaikoKeisan.Bind();
            BindCombo();
           
        }


        public void BindCombo()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("ID");
            dt.Columns.Add("Name");
            dt.Rows.Add("-1", string.Empty);
            dt.Rows.Add("0", "WEB在庫のみ");
            dt.Rows.Add("1", "豊中在庫数+Web通販在庫数");
           

            DataGridViewComboBoxColumn col = (DataGridViewComboBoxColumn)gdvPSKS0119M.Columns["colZaikoKeisan"];
            col.DataPropertyName = "ZaikoKeisan";
            col.DisplayMember = "Name";
            col.ValueMember = "ID";
            ((DataGridViewComboBoxColumn)gdvPSKS0119M.Columns["colZaikoKeisan"]).DataSource = dt;
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            
            OpenFileDialog op = new OpenFileDialog();
            op.InitialDirectory = @"C:\";
            op.RestoreDirectory = true;
            op.Filter = "Excel Files|*.xls;*.xlsx;";
            if (op.ShowDialog() == DialogResult.OK)
            {
                txtFilePath.Text = op.FileName;
            }
        }

        /// <summary>
        /// override base form's virtual function
        /// </summary>
        /// <param name="btnIndex"></param>
        public override void functionClick(int btnIndex)
        {
            switch (btnIndex)
            {
                case 1:
                    break;
                case 4:
                    DataDelete();
                    txtFilePath.Focus();
                    break;

                case 6:
                    Clear();
                    txtFilePath.Focus();
                    break;
                case 11:
                    ExcelExport();
                    break;

                case 12:
                   F12();
                   break;

            }
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            resultAdminCD = string.Empty;
            ImportData();
        }


        private void Clear()
        {
            txtFilePath.Text = string.Empty;
            cboTaisho.SelectedValue = -1;
            cboShiiresaki.SelectedValue = -1;
            cboBrand.SelectedValue = -1;
            txtRefNo.Text = string.Empty;
            cboZaikoKeisan.SelectedValue =-1;
            cboTaisho1.SelectedValue = -1;
            gdvPSKS0119M.DataSource = null;
            dtGridData = dtSearch = null;

        }

        private void ExcelExport()
        {
            if (dtSearch != null || dtGridData != null)  
            {
                if (DSP_MSG("Q205", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty).Equals("1"))
                {
                    DataTable dtExport = dtSearch != null && dtSearch.Rows.Count > 0 ? dtSearch.Copy() : dtGridData.Copy();
                    dtExport = ChangeDataColumnName(dtExport);
                    SaveFileDialog savedialog = new SaveFileDialog();
                    savedialog.Filter = "Excel Files|*.xlsx;";
                    savedialog.Title = "Save";
                    savedialog.FileName = "ExportFile";
                    savedialog.InitialDirectory = @"C:\";
                    savedialog.RestoreDirectory = true;
                    if (savedialog.ShowDialog() == DialogResult.OK)
                    {
                        if (Path.GetExtension(savedialog.FileName).Contains(".xlsx"))
                        {
                            Microsoft.Office.Interop.Excel._Application excel = new Microsoft.Office.Interop.Excel.Application();
                            Microsoft.Office.Interop.Excel._Workbook workbook = excel.Workbooks.Add(Type.Missing);
                            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;

                            worksheet = workbook.ActiveSheet;
                            worksheet.Name = "ExportedFromDatGrid";

                            using (XLWorkbook wb = new XLWorkbook())
                            {
                                wb.Worksheets.Add(dtExport, "test");

                                wb.SaveAs(savedialog.FileName);
                                DSP_MSG("I203", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);  //Export Successful


                            }
                        }

                    }
                }
            }
        }

        public void ImportData()
        {
            dtSearch=null;
            if (!string.IsNullOrWhiteSpace(resultAdminCD))
            {
                dtGridData = psks0119mbl.SelectDataByAdminCD(resultAdminCD);
                if (dtGridData != null)
                    gdvPSKS0119M.DataSource = dtGridData;
            }
            else
            {
                if (!string.IsNullOrWhiteSpace(txtFilePath.Text))
                {
                    String[] colName = { "メーカー商品CD", "JANCD", "基準数", "在庫計算", "指示書番号", "備考" };
                    string filePath = txtFilePath.Text;
                    if ((Have_opened_Excel(filePath)))  //  --- for Checking the imported excel are opened or not?
                    {
                        DataTable dtImport = new DataTable();

                        string ext = Path.GetExtension(filePath);
                        if (ext.Equals(".xlsx") || ext.Equals(".xls"))
                        {

                            dtImport = ExcelToDatatable(filePath);
                            if (dtImport != null)
                            {
                                if (CheckColumn(colName, dtImport)) // for checking column name from imported file
                                {
                                    BindCombo();
                                    dtImport = ChangeColName(dtImport);
                                    DataRow[] drNull = dtImport.Select("MakerShohin IS NULL and JANCD IS NULL");
                                    DataRow[] drMaker = dtImport.Select("MakerShohin IS NOT NULL ");
                                    DataRow[] drJanCD = dtImport.Select("JANCD IS NOT NULL ");

                                    if (drNull.Count() > 0 || (drMaker.Count() > 0 && drMaker.Count() != dtImport.Rows.Count) || (drJanCD.Count() > 0 && drJanCD.Count() != dtImport.Rows.Count))
                                    {
                                        DSP_MSG("E144", "名前", string.Empty, string.Empty, string.Empty, string.Empty);
                                        gdvPSKS0119M.DataSource = null;
                                    }

                                    else if (psks0119mbl.CheckMakerJanCD(dtImport))
                                    {
                                        DSP_MSG("E144", "名前", string.Empty, string.Empty, string.Empty, string.Empty);
                                        gdvPSKS0119M.DataSource = null;
                                    }

                                    else
                                    {
                                        if (DSP_MSG("Q206", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty).Equals("1"))
                                        {

                                            if (!string.IsNullOrWhiteSpace(dtImport.Rows[0]["MakerShohin"].ToString()) && string.IsNullOrWhiteSpace(dtImport.Rows[0]["JANCD"].ToString()))
                                                dtGridData = psks0119mbl.SelectData(dtImport, "1");

                                            else if (string.IsNullOrWhiteSpace(dtImport.Rows[0]["MakerShohin"].ToString()) && !string.IsNullOrWhiteSpace(dtImport.Rows[0]["JANCD"].ToString()))
                                                dtGridData = psks0119mbl.SelectData(dtImport, "2");

                                            else if (!string.IsNullOrWhiteSpace(dtImport.Rows[0]["MakerShohin"].ToString()) && !string.IsNullOrWhiteSpace(dtImport.Rows[0]["JANCD"].ToString()))
                                                dtGridData = psks0119mbl.SelectData(dtImport, "3");

                                            if (dtGridData != null)
                                                gdvPSKS0119M.DataSource = dtGridData;

                                        }
                                    }
                                }
                                else
                                {
                                    DSP_MSG("E137", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                                    txtFilePath.Focus();
                                }
                            }

                            //System.IO.File.Delete(filePath);
                        }

                    }
                    else
                    {
                        DSP_MSG("E122", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);   //Some of the imported files are being used by another process. please close those files to be able to import
                        txtFilePath.Focus();
                    }
                }
                else
                {
                    DSP_MSG("E122", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                    gdvPSKS0119M.DataSource = null;
                    txtFilePath.Focus();
                }
            }
        }

        private void DataDelete()
        {
            gdvPSKS0119M.EndEdit();
             adminCD = string.Empty;
             if (gdvPSKS0119M.Rows.Count > 0)
             {
                 
                 if (ArradminCD.Count > 0)
                 {

                     adminCD = string.Join(",", (string[])ArradminCD.ToArray(Type.GetType("System.String")));

                 }
                 if (string.IsNullOrWhiteSpace(adminCD))
                 {
                     DSP_MSG("E805", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);    //Please! Check the Checkbox
                 }
                 else
                 {
                     mzke = new M_Zaiko_Kanri_Entity
                     {
                         OperatorCD = loginInfo.OperatorCode,
                         AdminCD = adminCD,
                     };
                     lle = new L_Log_Entity
                     {
                         Program = "PSKS0119M",
                         PC = loginInfo.PcName,
                         OperateMode = "Delete",

                     };
                     if (psks0119mbl.M_Zaiko_RenkanriDelete(mzke, lle))
                     {

                         DSP_MSG("I102", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                         Clear();
                         txtFilePath.Focus();
                     }
                 }
             }
        }


        protected Boolean CheckColumn(String[] colName, DataTable dt)
        {
            DataColumnCollection col = dt.Columns;
            for (int i = 0; i < colName.Length; i++)
            {
                if (!col.Contains(colName[i]))
                    return false;
            }
            return true;
            
        }

        public void F12()
        {

            if(ErrorCheck())
            {
                adminCD = string.Empty;
                if (gdvPSKS0119M.Rows.Count > 0)
                {
                   if ((DSP_MSG("Q101", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty).Equals("1")))
                    {
                        DataTable dtUpdate = new DataTable();
                        dtUpdate.Columns.Add("OperatorCD");
                        dtUpdate.Columns.Add("AdminCD");
                        dtUpdate.Columns.Add("HanbaiShohinCD");
                        dtUpdate.Columns.Add("JanCD");
                        dtUpdate.Columns.Add("vc_maker_shohin");
                        dtUpdate.Columns.Add("taisho");
                        dtUpdate.Columns.Add("kijunsu");
                        dtUpdate.Columns.Add("Zaiko_keisan");
                        dtUpdate.Columns.Add("shijisho_CD");
                        dtUpdate.Columns.Add("Memo");

                        foreach (DataGridViewRow row in gdvPSKS0119M.Rows)
                        {
                            DataRow dtrow = dtUpdate.NewRow();
                            dtrow["OperatorCD"] = loginInfo.OperatorCode;
                            dtrow["AdminCD"] = row.Cells["colAdminCD"].Value.ToString();
                            dtrow["HanbaiShohinCD"] = row.Cells["colSKUCD"].Value.ToString();
                            dtrow["JanCD"] = row.Cells["colJANCD"].Value.ToString();
                            dtrow["vc_maker_shohin"] = row.Cells["colMakerCD"].Value.ToString();
                            dtrow["taisho"] = row.Cells["colTaisho"].Value.ToString()=="無効"? "0": "1";
                            dtrow["kijunsu"] =row.Cells["colRefNo"].Value.ToString();

                            DataGridViewComboBoxCell cmb = (DataGridViewComboBoxCell)row.Cells["colZaikoKeisan"];
                            dtrow["Zaiko_keisan"] = cmb.Value.ToString()=="-1" ? string.Empty :cmb.Value.ToString();

                            dtrow["shijisho_CD"] = row.Cells["colInstructionNo"].Value.ToString();
                            dtrow["Memo"] = row.Cells["colRemark"].Value.ToString();
                            dtUpdate.Rows.Add(dtrow);

                        }
                        mzke = new M_Zaiko_Kanri_Entity
                        {
                            dt=dtUpdate,
                          
                        };
                        lle = new L_Log_Entity
                        {
                            Program = "PSKS0119M",
                            PC = loginInfo.PcName,
                            OperateMode = "InsertUpdate",

                        };

                        if (psks0119mbl.M_Zaiko_KanriInsertUpdate(mzke, lle))
                        {
                            DSP_MSG("I101", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                            //Clear();
                            txtFilePath.Focus();
                        }
                    }

                }

            }
        }

        private L_Log_Entity GetLogData()
        {
            lle = new L_Log_Entity();
            lle.InsertOperator = loginInfo.OperatorCode;
            lle.Program = "PSKS0119M";
            lle.PC = loginInfo.PcName;
            lle.OperateMode = loginInfo.StartMode;
            //lle.KeyItem = ucSupplier.UC_Code + " " + lblKakuteiMode.Text;

            return lle;
        }

        public bool Have_opened_Excel(string filepath)
        {

            try
            {
                var f = File.OpenRead(filepath);
                f.Close();
            }
            catch
            {
                return false;
            }

                    return true;
        }

        private bool ErrorCheck()
        {
            foreach (DataGridViewRow dr in gdvPSKS0119M.Rows)
            {
                if (string.IsNullOrWhiteSpace(dr.Cells["colRefNo"].Value.ToString()))
                {
                    dr.Cells["colRefNo"].Value = 0;
                }

                DataGridViewComboBoxCell cmb = (DataGridViewComboBoxCell)dr.Cells["colZaikoKeisan"];
                if (cmb.Value.ToString()== "-1" || string.IsNullOrWhiteSpace(cmb.Value.ToString()))
                {
                    DSP_MSG("E102", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                    gdvPSKS0119M.CurrentCell = gdvPSKS0119M[dr.Cells["colZaikoKeisan"].ColumnIndex, dr.Index];
                    return false;
                }
                
            }

            return true;
        }

        protected DataTable ExcelToDatatable(string filePath)
        {
            FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read);

            string ext = Path.GetExtension(filePath);
            IExcelDataReader excelReader;
            if (ext.Equals(".xls"))
                //1. Reading from a binary Excel file ('97-2003 format; *.xls)
                excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
            else if (ext.Equals(".xlsx"))
                //2. Reading from a OpenXml Excel file (2007 format; *.xlsx)
                excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            else
                //2. Reading from a OpenXml Excel file (2007 format; *.xlsx) 
                excelReader = ExcelReaderFactory.CreateCsvReader(stream, null);

            //3. DataSet - The result of each spreadsheet will be created in the result.Tables
            bool useHeaderRow = true;
            
            DataSet result = excelReader.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = useHeaderRow,
                }
            });


            excelReader.Close();
            return result.Tables[0];
        }

        protected DataTable ChangeColName(DataTable dtImport)
        {
            dtImport.Columns["メーカー商品CD"].ColumnName = "MakerShohin";
            dtImport.Columns["基準数"].ColumnName = "RefNo";
            dtImport.Columns["在庫計算"].ColumnName = "ZaikoKeisan";
            dtImport.Columns["指示書番号"].ColumnName = "InstructionNo";
            dtImport.Columns["備考"].ColumnName = "Remark";
            return dtImport;
        }

        protected DataTable ChangeDataColumnName(DataTable dt)
        {
            dt.Columns["taisho"].ColumnName = "対象";
            dt.Columns["vc_maker_shohin"].ColumnName = "メーカー商品CD";
            dt.Columns["vm_kihon_shohin"].ColumnName = "商品名";
            dt.Columns["vm_color"].ColumnName = "カラー";
            dt.Columns["vm_size"].ColumnName = "サイズ";
            dt.Columns["RefNo"].ColumnName = "基準数";
            dt.Columns["ZaikoKeisan"].ColumnName = "在庫計算";
            dt.Columns["VM_SHIIRESAKI"].ColumnName = "仕入先名";
            dt.Columns["vm_brand"].ColumnName = "ブランド名";
            dt.Columns["InstructionNo"].ColumnName = "指示書番号";
            dt.Columns["Remark"].ColumnName = "備考";

            dt.Columns.RemoveAt(2);
            return dt;
        }

        
        private void btnAllCheck_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in gdvPSKS0119M.Rows)
            {
                DataGridViewCheckBoxCell chk = (DataGridViewCheckBoxCell)row.Cells[1];
                chk.Value = true; //because chk.Value is initialy null
                ArradminCD.Add(row.Cells["colAdminCD"].Value.ToString());
            }
        }

        private void btnAllUncheck_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in gdvPSKS0119M.Rows)
            {
                DataGridViewCheckBoxCell chk = (DataGridViewCheckBoxCell)row.Cells[1];
                chk.Value = false; //because chk.Value is initialy null
                ArradminCD.Remove(row.Cells["colAdminCD"].Value.ToString());
            }
        }

        private void btnSetting_Click(object sender, EventArgs e)
        {
            adminCD = string.Empty;
            if (gdvPSKS0119M.Rows.Count > 0)
            {
                if (txtRefNo.Text.Length <= 3)
                {

                    if (ArradminCD.Count > 0)
                        {
                            adminCD = string.Join(",", (string[])ArradminCD.ToArray(Type.GetType("System.String")));

                        }

                    if (string.IsNullOrWhiteSpace(adminCD))
                    {
                        DSP_MSG("E805", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);  //Please! Check the Checkbox
                    }
                    else
                    {

                        mzke = new M_Zaiko_Kanri_Entity
                        {
                            OperatorCD = loginInfo.OperatorCode,
                            AdminCD = adminCD,
                            taisho = cboTaisho1.SelectedValue.ToString(),
                            kijunsu = string.IsNullOrWhiteSpace(txtRefNo.Text) ? "-1" : txtRefNo.Text,
                            Zaiko_keisan = string.IsNullOrWhiteSpace(cboZaikoKeisan.SelectedValue.ToString()) ? "-1" : cboZaikoKeisan.SelectedValue.ToString(),

                        };
                        //if (psks0119mbl.M_Zaiko_Kanri_Setting(mzke))
                        //{
                            foreach (DataGridViewRow row in gdvPSKS0119M.Rows)
                            {
                                DataGridViewCheckBoxCell chk1 = row.Cells["colChk"] as DataGridViewCheckBoxCell;
                                if(Convert.ToBoolean(chk1.Value)==true)
                                {
                                    string taisho1 = mzke.taisho!="-1" && mzke.taisho == "0" ? "無効" : "有効";
                                    row.Cells["colTaisho"].Value = mzke.taisho == "-1" ? row.Cells["colTaisho"].Value.ToString() : taisho1; 
                                    row.Cells["colRefNo"].Value = mzke.kijunsu == "-1" ? row.Cells["colRefNo"].Value.ToString() : mzke.kijunsu;

                                    DataGridViewComboBoxCell cmb = (DataGridViewComboBoxCell)row.Cells["colZaikoKeisan"];
                                    cmb.Value = mzke.Zaiko_keisan == "-1" ? cmb.Value.ToString() : mzke.Zaiko_keisan;
                                }
                                chk1.Value = false;
                                ArradminCD.Remove(row.Cells["colAdminCD"].Value.ToString());
                            }
                            //DSP_MSG("I101", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                            txtFilePath.Focus();
                       // }
                    }
                }
                else
                {
                    DSP_MSG("E143", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);   //Character count is over!!
                    txtRefNo.Focus();
                }
              

            }
        }

        private void btnItemSearch_Click(object sender, EventArgs e)
        {
            FrmPSKS0110S_Search frmsks0110s = new FrmPSKS0110S_Search();
            frmsks0110s.ShowDialog();
            resultAdminCD = frmsks0110s.result;
            if(!string.IsNullOrWhiteSpace(resultAdminCD))
                ImportData();
        }

        private void gdvPSKS0119M_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
           
            if(gdvPSKS0119M.Rows[e.RowIndex].Cells["colTaisho"].Value.ToString()== "無効")
                gdvPSKS0119M.Rows[e.RowIndex].Cells["colTaisho"].Style.BackColor = System.Drawing.Color.Red;
            else
                gdvPSKS0119M.Rows[e.RowIndex].Cells["colTaisho"].Style.BackColor = System.Drawing.Color.FromArgb(217, 217, 217); ;
          

        }

        private void btnDisplay_Click(object sender, EventArgs e)
        {
            Search();
           
        }

        protected void Search()
        {
            string taisho = cboTaisho.SelectedValue.ToString() == "-1" ? string.Empty : cboTaisho.SelectedValue.ToString();
            string vc_shiiresaki = cboShiiresaki.SelectedValue.ToString() == "-1" ? string.Empty : cboShiiresaki.SelectedValue.ToString();
            string vc_brand = cboBrand.SelectedValue.ToString() == "-1" ? string.Empty : cboBrand.SelectedValue.ToString();

            dtGridData = null;
            ArradminCD.Clear();
            dtSearch = psks0119mbl.SelectDataByMakerCD(taisho, vc_shiiresaki, vc_brand,loginInfo.OperatorCode);
            if (dtSearch != null)
            {
                BindCombo();
                gdvPSKS0119M.DataSource = dtSearch;
            }
        }

        private void gdvPSKS0119M_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == gdvPSKS0119M.Columns["colRemark"].Index)
                gdvPSKS0119M.ImeMode = System.Windows.Forms.ImeMode.Hiragana;

            
        }

        private void gdvPSKS0119M_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            Maintained_CheckClick(sender, e);
        }

        protected void Maintained_CheckClick(object sender, DataGridViewCellEventArgs e)
        {

            if (e.ColumnIndex > 0 && e.RowIndex >= 0)
            {
                if ((sender as DataGridView).CurrentCell is DataGridViewCheckBoxCell)
                {
                    if ((Convert.ToBoolean(gdvPSKS0119M.Rows[e.RowIndex].Cells[e.ColumnIndex].EditedFormattedValue) == true))
                    {
                        DataGridViewCheckBoxCell chk1 = gdvPSKS0119M.Rows[e.RowIndex].Cells["colChk"] as DataGridViewCheckBoxCell;
                        ArradminCD.Add(gdvPSKS0119M.Rows[e.RowIndex].Cells["colAdminCD"].Value.ToString());

                    }
                    else
                    {
                        ArradminCD.Remove(gdvPSKS0119M.Rows[e.RowIndex].Cells["colAdminCD"].Value.ToString());
                    }
                }

            }
        }

        private void frmPSKS0119M_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F10)
                Search();
        }
        
    }
}
