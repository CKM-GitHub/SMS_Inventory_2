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
using System.IO;
using System.Collections;
using ExcelDataReader;
using System.Threading;

namespace SMS.Inventory
{
    public partial class frmPSKS0116I : SMS_Base
    {
        Login_Info loginInfo;
        M_MultiPorpose_Entity mme;
        T_MakerInportRireki_Entity tmire;
        T_MakerZaiko_Entity tmze;
        T_MakerInport_Entity tmie;
        M_SiiresakiZaiko_Entity msze;
        PSKS0103I_BL psks0103ibl;
        PSKS0116I_BL pskso116ibl;

        int SearchMode = 1;
        
        string InportSEQ = string.Empty;
        bool Chk_flg = false;
        string folderPath = string.Empty;
        string FileName = string.Empty;
        static string ErrChkFlg = "0";
        int Delete_row = 0;

        public frmPSKS0116I()
        {
            InitializeComponent();
        }
        
        public frmPSKS0116I(Login_Info loginInfo)
            : base(loginInfo)
        {
            InitializeComponent();
            this.loginInfo = loginInfo;          
        }

        private void frmPSKS0103I_Load(object sender, EventArgs e)
        {
            FunctionButtonDisabled(2);
            FunctionButtonDisabled(3);
            FunctionButtonDisabled(4);
            FunctionButtonDisabled(5);
            FunctionButtonDisabled(7);
            FunctionButtonDisabled(8);
            FunctionButtonDisabled(10);
            FunctionButtonDisabled(11);

            pskso116ibl = new PSKS0116I_BL();
            psks0103ibl = new PSKS0103I_BL();
            ucSupplier.GotFocus += new System.EventHandler(ucSupplier_GotFocus);

            FormName = "メーカー在庫取込";

            lblMode.Visible = true;
            lblMode.Text = "先に通常取込を実施してください";
            lblMode.BackColor = Color.FromArgb(255, 102, 255);
            lblMode.Width = 250;
            lblMode.TextAlign = ContentAlignment.MiddleCenter;

            btnImport.Enabled = false;
            FunctionButtonDisabled(12);

            mme = new M_MultiPorpose_Entity();
            mme.fields = "Char1,Num1";
            mme.tableName = "M_MultiPorpose";
            mme.condition = "Where ID=101 And [Key]='1'";

            DataTable dt = psks0103ibl.M_MultiPorpose_DynamicSelect(mme);
            if (dt.Rows.Count > 0)
            {
                lblServerPath.Text = "サーバーの" + dt.Rows[0]["Char1"].ToString() + "フォルダーに保存されているファイルを取込します。";
                lblHistory.Text = dt.Rows[0]["Num1"] + "日間の履歴を保持しています";
                folderPath = dt.Rows[0]["Char1"].ToString();
            }

            dgvImportRireki.DisabledColumn("colNo1,colInportDateTime,colDataMoto,colInportFile,colInportOperator,colInportSu,colErrorSu");
            dgvFileDetail.DisabledColumn("*");

            tmire = new T_MakerInportRireki_Entity();
            DataTable dtInport = psks0103ibl.T_MakerInportRireki_SELECT(tmire);
            dgvImportRireki.DataSource = dtInport;           
        }

        private void CreateDataTablefor_dgvImportRireki()
        {
            DataTable dt = new DataTable();

            dt.Columns.Add("No", typeof(string));
            dt.Columns.Add("chk1", typeof(string));
            dt.Columns.Add("chk2", typeof(string));
            dt.Columns.Add("chk3", typeof(string));
            dt.Columns.Add("InsertDateTime", typeof(string));
            dt.Columns.Add("VM_SHIIRESAKI", typeof(string));
            dt.Columns.Add("InportFile", typeof(string));
            dt.Columns.Add("OperatorName", typeof(string));
            dt.Columns.Add("InportSu", typeof(string));
            dt.Columns.Add("ErrorKBN", typeof(string));
            dt.Columns.Add("InportSEQ", typeof(string));
            dgvImportRireki.DataSource = dt;        
        }

        private void btnFolderConfirm_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(folderPath))
            {
                if (!Directory.Exists(folderPath))
                {
                    Directory.CreateDirectory(folderPath);
                    System.Diagnostics.Process.Start(folderPath);
                }
                else
                {
                    System.Diagnostics.Process.Start(folderPath);
                }
            }
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            F12();
        }

        private void F12()
        {
            if (ErrorCheck())//1             
            {
                if (DSP_MSG("Q101", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty).Equals("1"))
                {
                    ArrayList FilePath = getFilePath();//Modify2018-11-09
                    tmze = new T_MakerZaiko_Entity();
                    tmze.MakerCD = ucSupplier.UC_Code;
                    tmze.InsertOperator = loginInfo.OperatorCode;

                    //PopUp.PSKS0103I_Import import = new PopUp.PSKS0103I_Import(FilePath, tmze,"1");
                    //import.ShowDialog();

                    InportRireki_Select();
                }

            }
            else if (ErrChkFlg == "3")
            {
                DeleteBy_DeleteCheck();
            }
        }

        private void InportRireki_Select() //【取込履歴】gridview Bind
        {
            tmire = new T_MakerInportRireki_Entity();
            tmire.DataSourseMakerCD = ucSupplier.UC_Code;
            DataTable dtInportRireki = psks0103ibl.T_MakerInportRireki_SELECT(tmire);
            dgvImportRireki.DataSource = dtInportRireki;
        }

        private void DeleteBy_DeleteCheck()
        {
            tmie = new T_MakerInport_Entity();
            InportSEQ =string.Empty;
            foreach (DataGridViewRow row in dgvImportRireki.Rows)
            {
                if (Convert.ToBoolean(row.Cells["chk3"].Value) == true)
                {
                    if (InportSEQ != string.Empty)
                        InportSEQ += ",";

                    InportSEQ += row.Cells["colInportSEQ"].Value.ToString();
                }
            }
            tmie.InportSEQ = InportSEQ;
            tmie.InsertOperator = loginInfo.OperatorCode;
            tmie.InsertDateTime = System.DateTime.Now.ToString();
            psks0103ibl.PSKS0103I_DELETE(tmie);
            InportRireki_Select();
        }

        protected DataTable ExcelToDatatableCSV(string filePath)
        {
            FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read);

            string ext = Path.GetExtension(filePath);
            IExcelDataReader excelReader;
            if (ext.Equals(".csv"))
                //1. Reading from a binary Excel file ('97-2003 format; *.xls)
                excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
            else if (ext.Equals(".xlsx"))
                //2. Reading from a OpenXml Excel file (2007 format; *.xlsx)
                excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            else
                //2. Reading from a OpenXml Excel file (2007 format; *.xlsx)
                excelReader = ExcelReaderFactory.CreateCsvReader(stream);

            //3. DataSet - The result of each spreadsheet will be created in the result.Tables

            DataSet result = excelReader.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = true
                }
            });


            excelReader.Close();
            return result.Tables[0];
        }

        private bool ErrorCheck()
        {
            if (!string.IsNullOrWhiteSpace(ucSupplier.UC_Code))
            {
                if (!ucSupplier.SelectData())
                {
                    DSP_MSG("E101", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                    ucSupplier.SetFocus();
                    //btnImport.Enabled = false;
                    return false;
                }
                else if (!psks0103ibl.MakerBrandSelect("MakerCD", "M_MakerBrand", "Where MakerCD='" + ucSupplier.UC_Code + "'"))
                {
                    DSP_MSG("E101", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty); 
                    ucSupplier.SetFocus();
                    //btnImport.Enabled = false;
                    return false;
                }
            }
            else
            {
                //DSP_MSG("E102", "名前", string.Empty, string.Empty, string.Empty, string.Empty);
                ucSupplier.SetFocus();
                //btnImport.Enabled = false;
                return false;
            }

            ErrChkFlg = "0";
            foreach (DataGridViewRow row1 in dgvImportRireki.Rows)
            {
                DataGridViewCheckBoxCell chk1 = row1.Cells["chk1"] as DataGridViewCheckBoxCell;
                DataGridViewCheckBoxCell chk2 = row1.Cells["chk2"] as DataGridViewCheckBoxCell;
                DataGridViewCheckBoxCell chk3 = row1.Cells["chk3"] as DataGridViewCheckBoxCell;

                if (chk1.Value == null) chk1.Value = false;
                if (chk2.Value == null) chk2.Value = false;
                if (chk3.Value == null) chk3.Value = false;

                if ((bool)chk1.Value != false && (bool)chk2.Value != false)// check「結果」「エラー」
                {
                    DSP_MSG("E126", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                    return false;
                }
                else if ((bool)chk3.Value == false)//Check File Type and Folder path
                {
                    //if (!File.Exists(getFilePath()))
                    if (getFilePath().Count <= 0)
                    {
                        DSP_MSG("E127", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                        return false;
                    }
                }              
                else
                {
                    ErrChkFlg = "3";
                    Delete_row = row1.Index;
                    return false;
                }               
            }

           
            if(string.IsNullOrWhiteSpace(lblFormatCD.Text))
            {
                MessageBox.Show("フォーマット Doesn't Exists", "メッセージ");
                return false;
            }
            if (psks0103ibl.JANCD_PositionCheck(ucSupplier.UC_Code).Equals("0"))
            {
                MessageBox.Show("JAN Position Not Found!", "メッセージ");
                return false;
            }

            return true;
        }

        private ArrayList getFilePath()
        {
            ArrayList arr = new ArrayList();

            string type = "csv";
            if (lblFileType.Text == "Excel")
            {
                type = "xlsx";
            }
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            DirectoryInfo d = new DirectoryInfo(folderPath);
            FileInfo[] Files = d.GetFiles("*." + type);
            foreach (FileInfo file in Files)
            {
                if (file.Name.StartsWith(FileName))
                {
                    arr.Add(file.FullName);
                }
            }
            return arr;
        }

        private void ucSupplier_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                #region***BugNo1***
                if (string.IsNullOrWhiteSpace(ucSupplier.UC_Code))
                {
                    DSP_MSG("E102", "名前", string.Empty, string.Empty, string.Empty, string.Empty);
                    ucSupplier.SetFocus();
                    return;
                }
                #endregion

                if (ucSupplier.SelectData())
                {
                    msze = new M_SiiresakiZaiko_Entity();
                    msze.SiiresakiCD = ucSupplier.UC_Code;

                    DataTable dt = pskso116ibl.M_SiiresakiZaiko_Select(msze);
                    if (dt.Rows.Count > 0)
                    {
                        lblFormatCD.Text = dt.Rows[0]["PatternCD"].ToString();
                        lblFormatName.Text = dt.Rows[0]["PatternName"].ToString();
                        lblTroikomiFile.Text = "ファイル名が" + dt.Rows[0]["FileNameWord"].ToString() + "から始まるファイルを取込する。";
                        lblFileType.Text = dt.Rows[0]["FileKBN"].ToString();
                        lblHiddenKubun.Text = dt.Rows[0]["FirstRecordKBN"].ToString();
                        FileName = dt.Rows[0]["FileNameWord"].ToString();
                        btnImport.Enabled = true;
                        FunctionButtonEnabled(12);
                    }
                    else
                    {
                         lblFormatCD.Text = "";
                         lblFormatName.Text = "";
                         lblTroikomiFile.Text = "ファイル名がxxxから始まるファイルを取込する。";
                         lblFileType.Text = ". . .";
                         btnImport.Enabled = false;
                         FunctionButtonDisabled(12);
                    }
                    InportRireki_Select();
                    //ktp 2018/12/31 no need to disable on any condition
                    ucSupplier.Disabled();
                }
                else
                {
                    DSP_MSG("E101", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                    btnImport.Enabled = false;
                    FunctionButtonDisabled(12);
                    dgvImportRireki.DataSource = null;
                }
                dgvFileDetail.DataSource = null;
                Chk_flg = false;
                InportSEQ = string.Empty;
            }
            //else
            //{
            //    lblFormatCD.Text = "";
            //    lblFormatName.Text = "";
            //    lblTroikomiFile.Text = "ファイル名がxxxから始まるファイルを取込する。";
            //    lblFileType.Text = ". . .";
            //    btnImport.Enabled = false;
            //    FunctionButtonDisabled(12);
            //    Chk_flg = false;
            //}
        }

        private void dgvImportRireki_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)//must not header row
            {
                if ((sender as DataGridView).CurrentCell is DataGridViewCheckBoxCell)
                {
                    ErrChkFlg = "0";
                    if ((Convert.ToBoolean(dgvImportRireki.Rows[e.RowIndex].Cells[e.ColumnIndex].EditedFormattedValue) == true))
                    {
                        
                        foreach (DataGridViewRow row1 in dgvImportRireki.Rows)
                        {
                            DataGridViewCheckBoxCell chk1 = row1.Cells["chk1"] as DataGridViewCheckBoxCell;
                            DataGridViewCheckBoxCell chk2 = row1.Cells["chk2"] as DataGridViewCheckBoxCell;
                            DataGridViewCheckBoxCell chk3 = row1.Cells["chk3"] as DataGridViewCheckBoxCell;
                            chk1.Value = chk1.FalseValue;
                            chk2.Value = chk2.FalseValue;
                            if (dgvImportRireki.CurrentCell.ColumnIndex != dgvImportRireki.Columns["chk3"].Index)
                            {
                                chk3.Value = chk3.FalseValue;
                                if (row1.Index % 2 == 0)
                                {
                                    row1.DefaultCellStyle.BackColor = Color.White;
                                }
                                else
                                {
                                    row1.DefaultCellStyle.BackColor = Color.FromArgb(221, 235, 247); 
                                }
                            }
                        }

                        dgvImportRireki.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = true;
                        dgvImportRireki.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Yellow;

                        if ((dgvImportRireki.Rows[e.RowIndex].Cells["chk1"].Selected == true) || (dgvImportRireki.Rows[e.RowIndex].Cells["chk2"].Selected == true))
                        {
                            Chk_flg = true;
                            ErrChkFlg = dgvImportRireki.Rows[e.RowIndex].Cells["chk2"].Selected == false ? "0" : "1";
                            InportSEQ = dgvImportRireki.Rows[e.RowIndex].Cells["colInportSEQ"].Value.ToString();
                        }
                        else if (dgvImportRireki.Rows[e.RowIndex].Cells["chk3"].Selected == true)
                        {
                           
                            ErrChkFlg = "3";
                          
                        }
                        else
                        {
                            Chk_flg = false;
                        }
                    }
                    else
                    {
                        InportSEQ = string.Empty;
                        Chk_flg = false;
                        FunctionButtonDisabled(12);
                        btnImport.Enabled = false;
                        dgvImportRireki.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.White;
                        dgvImportRireki.AllowUserToAddRows = false;
                    }

                    
                }
                else
                {
                    Chk_flg = false;

                }
            }
            if (ErrChkFlg == "3")
            {
                FunctionButtonEnabled(12);
                btnImport.Enabled = true;
            }
            
            
        }

        public async Task<bool> Detail()
        {
            Func<bool> function = new Func<bool>(() => ShowDetail());
            return await Task.Factory.StartNew<bool>(function);
        }

        private bool ShowDetail()
        {
            if (btnF11_Detail.InvokeRequired)
            {
                btnF11_Detail.Invoke(new MethodInvoker(delegate
                {
                DataTable dtmakerinput = new DataTable();
                dtmakerinput = psks0103ibl.T_MakerInport_Select(InportSEQ, ErrChkFlg);
                dgvFileDetail.DataSource = dtmakerinput;
                    }));
            }
            return true;
        }

        private async void F11()
        {
            if (Chk_flg)
            {
                WaitWnd.WaitWndFun waitForm = new WaitWnd.WaitWndFun("Searching.....Please Wait!");
                waitForm.Show(this);

                await Detail();

                waitForm.Close();
                btnF11_Detail.Focus();
            }
        }

        private void btnF11_Detail_Click(object sender, EventArgs e)
        {
            F11();
        }

        public override void functionClick(int btnIndex)
        {
            switch (btnIndex)
            {
                case 1:
                    break;
                case 6:
                    Clear();
                    //dgvImportRireki.Enabled = false;
                    //dgvFileDetail.Enabled = false;
                    //dgvImportRireki.DataSource = null;
                    dgvFileDetail.DataSource = null;
                    ucSupplier.SetFocus();
                    ucSupplier.Enable();
                    Chk_flg = false;
                    break;
               case 9:
                    F9();
                    break;
               case 12:
                    F12();
                    break;
            }
        }
        private void Clear()
        {
            ucSupplier.UC_Code = string.Empty;
            ucSupplier.UC_Name = string.Empty;
            tmire = new T_MakerInportRireki_Entity();
            DataTable dtInport = psks0103ibl.T_MakerInportRireki_SELECT(tmire);
            dgvImportRireki.DataSource = dtInport;
            lblTroikomiFile.Text = "ファイル名がXXXから始まるファイルを取込する。";
            lblFileType.Text = "...";
            lblHiddenKubun.Text = string.Empty;
            lblFormatCD.Text = string.Empty;
            lblFormatName.Text = string.Empty;
        }

        private void ucSupplier_Enter(object sender, EventArgs e)
        {
            SearchMode = 1;
            if (ucSupplier.UC_SearchEnable)
            {             
                FunctionButtonEnabled(9);
            }

            btnImport.Enabled = false;
            FunctionButtonDisabled(12);
        }

        private void F9()
        {
            if (SearchMode == 1)
                ucSupplier.ShowSearchForm();
        }

        private void frmPSKS0103I_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F11)
                F11();
        }

        private void ucSupplier_GotFocus(object sender, EventArgs e)
        {
            btnImport.Enabled = false;
        }
    }
}
