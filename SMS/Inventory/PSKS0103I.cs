using ExcelDataReader;
using SMS_BL;
using SMS_Entity;
using System;
using System.Collections;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using ElencySolutions.CsvHelper;
using System.Text;
using System.Diagnostics;

namespace SMS.Inventory
{
    public partial class frmPSKS0103I : SMS_Base
    {
        Login_Info loginInfo;
        M_MultiPorpose_Entity mme;
        T_MakerInportRireki_Entity tmire;
        T_MakerZaiko_Entity tmze;
        T_MakerInport_Entity tmie;
        M_SiiresakiZaiko_Entity msze;
        PSKS0103I_BL psks0103ibl;
        int SearchMode = 0;
        string InportSEQ = string.Empty;
        string folderPath = string.Empty;
        string FileName = string.Empty;
        string ImportFileName = string.Empty;
        bool Disable_Flag = false;
        static string ErrChkFlg = string.Empty;
        int Delete_row = 0;

        DataTable dtmakerinput;

        #region constructor
        public frmPSKS0103I()
        {
            InitializeComponent();
        }

        public frmPSKS0103I(Login_Info loginInfo)
            : base(loginInfo)
        {
            InitializeComponent();
            this.loginInfo = loginInfo;
        }
        #endregion

        /// <summary>Load
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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

            psks0103ibl = new PSKS0103I_BL();
            ucSupplier.GotFocus += new System.EventHandler(ucSupplier_GotFocus);

            FormName = "メーカー在庫取込";
            lblMode.Visible = false;

            if (!ucSupplier.Enabled)
            {
                btnImport.Enabled = false;
                FunctionButtonDisabled(12);
            }

            mme = new M_MultiPorpose_Entity();
            mme.fields = "Char1,Num1,Char5";
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
            dgvFileDetail.DisabledSorting();
            dgvImportRireki.DisabledSorting();

            tmire = new T_MakerInportRireki_Entity();
            DataTable dtInport = psks0103ibl.T_MakerInportRireki_SELECT(tmire);
            dgvImportRireki.DataSource = dtInport;
        }

        /// <summary>Bind FileType Combobox
        /// 1.Csv, 2.Excel
        /// </summary>
        private void BindFileType()
        {
            DataTable dtFileType = new DataTable();
            dtFileType.Columns.Add("TypeID");
            dtFileType.Columns.Add("Extension");

            dtFileType.Rows.Add("1", "CSV");
            dtFileType.Rows.Add("2", "EXCEL");

            cboFileType.DataSource = dtFileType;
            cboFileType.DisplayMember = "Extension";
            cboFileType.ValueMember = "TypeID";
        }

        /// <summary>Open Import Folder
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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

        /// <summary>Import File(F12)
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnImport_Click(object sender, EventArgs e)
        {
            F12();
        }

        /// <summary>Import
        /// 
        /// </summary>
        private void F12()
        {
            if (ErrChkFlg == "3")
            {
                Cursor.Current = Cursors.WaitCursor;
                DeleteBy_DeleteCheck();
                Cursor.Current = Cursors.Default;
            }
            else
            {
                ArrayList FilePath = getFilePath();
                if (ErrorCheck(FilePath))//1             
                {
                    if (DSP_MSG("Q101", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty).Equals("1"))
                    {
                        //Modify2018-11-09   
                        tmze = new T_MakerZaiko_Entity();
                        tmze.MakerCD = ucSupplier.UC_Code;
                        tmze.InsertOperator = loginInfo.OperatorCode;


                        PopUp.PSKS0103I_Import import = new PopUp.PSKS0103I_Import(FilePath, tmze, lblHiddenKubun.Text, cboPatternCD.Text);
                        import.ShowDialog();

                        InportRireki_Select();
                    }
                }
            }
        }

        #region ucSupplier
        /// <summary>handle enter key press
        /// select data when enter key press
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ucSupplier_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                getInfo();
            }
        }

        /// <summary>handle other control on focus
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ucSupplier_GotFocus(object sender, EventArgs e)
        {
            if (!ucSupplier.Enabled)
            {
                btnImport.Enabled = false;
            }
            else
            {
                if (Disable_Flag)
                {
                    btnImport.Enabled = false;
                    FunctionButtonDisabled(12);
                }
                else
                {
                    btnImport.Enabled = true; FunctionButtonEnabled(12);
                }
            }
        }

        /// <summary>set import data
        /// 
        /// </summary>
        private void getInfo()
        {
            if (ucSupplier.UC_Exists)
            {
                msze = new M_SiiresakiZaiko_Entity();
                msze.SiiresakiCD = ucSupplier.UC_Code;

                DataTable dt = psks0103ibl.M_SiiresakiZaiko_Select(msze);
              
                    if (dt.Rows.Count > 0)
                    {
                        cboPatternCD.DataSource = dt;
                        cboPatternCD.DisplayMember = "PatternCD";
                        cboPatternCD.ValueMember = "PatternName";

                        lblFormatCD.Text = dt.Rows[0]["PatternCD"].ToString();
                        lblFormatName.Text = dt.Rows[0]["PatternName"].ToString();
                        lblTroikomiFile.Text = "ファイル名が" + dt.Rows[0]["FileNameWord"].ToString() + "から始まるファイルを取込する。";

                        BindFileType();
                        if (dt.Rows[0]["FileKBN"].ToString().Equals("CSV")) /// 2 is for excel(xlsx) and 1 is for csv
                        {
                            cboFileType.SelectedValue = 1; // csv
                        }
                        else
                        {
                            cboFileType.SelectedValue = 2; // excel
                        }

                        lblHiddenKubun.Text = dt.Rows[0]["FirstRecordKBN"].ToString();
                        FileName = dt.Rows[0]["FileNameWord"].ToString();
                        if (!ucSupplier.Enabled)
                        {
                            btnImport.Enabled = false;
                            FunctionButtonDisabled(12);
                        }
                        else
                        {
                            btnImport.Enabled = true;
                            FunctionButtonEnabled(12);
                        }
                    }
                    else
                    {
                        lblFormatCD.Text = "";
                        lblFormatName.Text = "";
                        lblTroikomiFile.Text = "ファイル名がxxxから始まるファイルを取込する。";
                        lblFileType.Text = ". . .";
                        cboPatternCD.DataSource = null;
                        cboFileType.DataSource = null;
                        if (!ucSupplier.Enabled)
                        {
                            btnImport.Enabled = false;
                            FunctionButtonDisabled(12);
                        }
                    }
                    InportRireki_Select();
                    //ktp 2018/12/31 no need to disable on any condition
                    ucSupplier.Disabled();

            }
            else
            {
                if (!ucSupplier.Enabled)
                {
                    btnImport.Enabled = false;
                    FunctionButtonDisabled(12);
                }
            }
            dgvFileDetail.DataSource = null;
            InportSEQ = string.Empty;
        }

        #region handle Search Form Popup for Supplier
        /// <summary>
        /// set searchMode
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ucSupplier_Enter(object sender, EventArgs e)
        {
            SearchMode = 1;
            FunctionButtonEnabled(9);
        }

        /// <summary>
        /// reset searchMode
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ucSupplier_Leave(object sender, EventArgs e)
        {
            if (this.ActiveControl.Name != "btnF9")
            {
                FunctionButtonDisabled(9);
                SearchMode = 0;
            }
        }

        /// <summary>show supplier search form
        /// 
        /// </summary>
        private void F9()
        {
            if (SearchMode == 1)
                ucSupplier.ShowSearchForm();
        }
        #endregion
        #endregion 

        #region handle F11
        /// <summary> Search Imported data by import file
        ///  Added by Mr.Phyoe 2_14_2109
        /// </summary>
        protected void F11()
        {
            Cursor.Current = Cursors.WaitCursor;
            dtmakerinput = new DataTable();
            dtmakerinput = psks0103ibl.T_MakerInport_Select(InportSEQ, ErrChkFlg);
            dgvFileDetail.DataSource = dtmakerinput;
            if (dgvFileDetail.Rows.Count > 0)
            {
                foreach (DataGridViewRow dr in dgvFileDetail.Rows)
                {
                    if (dr.Cells["colKBN"].Value.ToString() == "1")
                    {
                        dr.DefaultCellStyle.ForeColor = System.Drawing.Color.Blue;

                    }
                }
            }
            Cursor.Current = Cursors.Default;
        }

        /// <summary>F11 button click
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnF11_Detail_Click(object sender, EventArgs e)
        {
            if(btnF11_Detail.Enabled)
                F11();
        }

        /// <summary>handel function key F11 press
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmPSKS0103I_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F11)
                F11();
        }
        #endregion

        /// <summary>Handle F1 to F12 Click
        /// 
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
                    dgvFileDetail.DataSource = null;
                    ucSupplier.Enable();
                    ucSupplier.SetFocus();
                    break;
                case 9:
                    F9();
                    break;
                case 12:
                    F12();
                    break;
            }
        }

        /// <summary>Reset Data
        /// 
        /// </summary>
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
            cboPatternCD.DataSource = null;
            cboFileType.DataSource = null;
            lblFormatName.Text = string.Empty;
        }

        #region dgvImportRireki
        /// <summary>//【取込履歴】gridview Bind
        /// 
        /// </summary>
        private void InportRireki_Select()
        {
            tmire = new T_MakerInportRireki_Entity();
            tmire.DataSourseMakerCD = ucSupplier.UC_Code;
            DataTable dtInportRireki = psks0103ibl.T_MakerInportRireki_SELECT(tmire);
            dgvImportRireki.DataSource = dtInportRireki;
        }

        /// <summary>handle checkbox check
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvImportRireki_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            Maintained_CheckClick(sender, e);
        }

        /// <summary> change cursor
        /// Make Hand Cursor inside the Cell and Cell Content
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvImportRireki_CellMouseEnter(object sender, DataGridViewCellEventArgs e) 
        {
            if (e.ColumnIndex > 0 && e.RowIndex >=0)
            {
                if ((sender as DataGridView).CurrentCell is DataGridViewCheckBoxCell)
                {
                    switch ((sender as DataGridView).Columns[e.ColumnIndex].Name)
                    {
                        case ("chk1"): (sender as DataGridView).Cursor = Cursors.Hand; break;
                        case ("chk2"): (sender as DataGridView).Cursor = Cursors.Hand; break;
                        case ("chk3"): (sender as DataGridView).Cursor = Cursors.Hand; break;
                        // || ("chk2") || ("chk3")
                        default: (sender as DataGridView).Cursor = Cursors.Default; break;
                    }
                }
                else
                {
                    (sender as DataGridView).Cursor = Cursors.Arrow;
                }
            }
            else
            {
                //(sender as DataGridView).Cursor = Cursors.Arrow;
            }
        }

        /// <summary>change cursor
        /// Make Arrow Outsided the GridView cell
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvImportRireki_CellMouseLeave(object sender, DataGridViewCellEventArgs e)   
        {
            if (e.RowIndex >= 0)//must not header row
            {
                if ((sender as DataGridView).CurrentCell is DataGridViewCheckBoxCell)
                {
                    (sender as DataGridView).Cursor = Cursors.Arrow;
                }
            }
        }

        /// <summary>handle checkbox check
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvImportRireki_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            Maintained_CheckClick(sender, e);
        }

        /// <summary>Create Datatable
        /// 
        /// </summary>
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

        protected void Maintained_CheckClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex > 0 && e.RowIndex >= 0)
            {
                if ((sender as DataGridView).CurrentCell is DataGridViewCheckBoxCell)
                {
                    ErrChkFlg = "";
                    if ((Convert.ToBoolean(dgvImportRireki.Rows[e.RowIndex].Cells[e.ColumnIndex].EditedFormattedValue) == true))
                    {
                        foreach (DataGridViewRow row1 in dgvImportRireki.Rows)
                        {
                            DataGridViewCheckBoxCell chk1 = row1.Cells["chk1"] as DataGridViewCheckBoxCell;
                            DataGridViewCheckBoxCell chk2 = row1.Cells["chk2"] as DataGridViewCheckBoxCell;
                            DataGridViewCheckBoxCell chk3 = row1.Cells["chk3"] as DataGridViewCheckBoxCell;

                            if (dgvImportRireki.CurrentCell.ColumnIndex == dgvImportRireki.Columns["chk1"].Index)
                            {
                                ErrChkFlg = "";
                                ImportFileName= dgvImportRireki.Rows[e.RowIndex].Cells["colInportFile"].Value.ToString();
                                InportSEQ = dgvImportRireki.Rows[e.RowIndex].Cells["colInportSEQ"].Value.ToString();
                                if (row1 == dgvImportRireki.CurrentRow)
                                {
                                    chk1.Value = chk1.TrueValue;
                                    chk2.Value = chk2.FalseValue;
                                    chk3.Value = chk3.FalseValue;
                                    dgvImportRireki.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Yellow;
                                }
                                else
                                {
                                    chk1.Value = chk1.FalseValue;
                                    chk2.Value = chk2.FalseValue;
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
                            else if (dgvImportRireki.CurrentCell.ColumnIndex == dgvImportRireki.Columns["chk2"].Index)
                            {
                                InportSEQ = dgvImportRireki.Rows[e.RowIndex].Cells["colInportSEQ"].Value.ToString();
                                ImportFileName = dgvImportRireki.Rows[e.RowIndex].Cells["colInportFile"].Value.ToString();
                                ErrChkFlg = "2";
                                if (row1 == dgvImportRireki.CurrentRow)
                                {
                                    chk2.Value = chk2.TrueValue;
                                    chk1.Value = chk1.FalseValue;
                                    chk3.Value = chk3.FalseValue;

                                    dgvImportRireki.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Yellow;
                                }
                                else
                                {
                                    chk1.Value = chk1.FalseValue;
                                    chk2.Value = chk2.FalseValue;
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
                            else
                            {
                                ErrChkFlg = "3";
                                ImportFileName = dgvImportRireki.Rows[e.RowIndex].Cells["colInportFile"].Value.ToString();
                                if (row1 == dgvImportRireki.CurrentRow)
                                {
                                    chk3.Value = chk3.TrueValue;
                                    chk1.Value = chk1.FalseValue;
                                    chk2.Value = chk2.FalseValue;

                                    dgvImportRireki.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Yellow;
                                }
                                else
                                {
                                    chk1.Value = chk1.FalseValue;
                                    chk2.Value = chk2.FalseValue;
                                    if (chk3.Value == chk3.FalseValue)
                                    {
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
                            }
                        }
                        btnImport.Enabled = false;
                        FunctionButtonDisabled(12);
                        btnF11_Detail.Enabled = true;
                    }
                    else
                    {
                        var flg = "0";
                        if (dgvImportRireki.CurrentRow.Index % 2 == 0)
                        {
                            dgvImportRireki.CurrentRow.DefaultCellStyle.BackColor = Color.White;
                        }
                        else
                        {
                            dgvImportRireki.CurrentRow.DefaultCellStyle.BackColor = Color.FromArgb(221, 235, 247);
                        }

                        InportSEQ = string.Empty;
                        if (dgvImportRireki.CurrentCell.ColumnIndex == 3)   // added by Phyoe gyi 03082019
                        {
                            foreach (DataGridViewRow gr in dgvImportRireki.Rows)
                            {
                                if (Convert.ToBoolean(gr.Cells["chk3"].EditedFormattedValue) == true)
                                {
                                    flg = "1";  //checked
                                    break;

                                }
                            }
                            if (flg == "1")
                            {
                                btnF11_Detail.Enabled = false;
                                FunctionButtonEnabled(12);
                                btnImport.Enabled = true;
                            }
                            else
                            {
                                btnImport.Enabled = false;
                                FunctionButtonDisabled(12);
                                btnF11_Detail.Enabled = true;
                            }
                        }
                    }
                    if (ErrChkFlg == "3")
                    {
                        FunctionButtonEnabled(12);
                        btnImport.Enabled = true;
                        btnF11_Detail.Enabled = false;
                    }
                }
            }
        }
        #endregion

        #region cboPattern
        private void cboPatternCD_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboPatternCD.SelectedValue != null)
                lblFormatName.Text = cboPatternCD.SelectedValue.ToString();
            else
                lblFormatName.Text = string.Empty;
        }

        private void cboPatternCD_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = true;
        }
        #endregion

        /// <summary>change file type
        /// Update File type to database
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cboFileType_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboFileType.SelectedValue != null && cboFileType.SelectedValue.ToString() != "System.Data.DataRowView")
            {
                msze = new M_SiiresakiZaiko_Entity();
                msze.SiiresakiCD = ucSupplier.UC_Code; 
                msze.FileKBN = cboFileType.SelectedValue.ToString();

                psks0103ibl.M_SiiresakiZaiko_Update(msze);
            }
        }

        /// <summary>delete import history
        /// 
        /// </summary>
        private void DeleteBy_DeleteCheck()
        {
            if (DSP_MSG("Q102", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty).Equals("1"))
            {
                tmie = new T_MakerInport_Entity();
                InportSEQ = string.Empty;
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
        }

        /// <summary>Chcek file is lock
        /// 
        /// </summary>
        /// <param name="fileArray"></param>
        /// <returns></returns>
        public bool Have_opened_Excel(ArrayList fileArray)
        {
            string destination = folderPath;
            foreach (string path in fileArray)
            {
                try
                {
                    var f = File.OpenRead(path);
                    f.Close();
                }
                catch
                {
                    return false;
                }
            }
            return true;
        }

        /// <summary>Error Check Before Import
        /// 
        /// </summary>
        /// <param name="FilePath">import file path</param>
        /// <returns></returns>
        private bool ErrorCheck(ArrayList FilePath)
        {
            if (!string.IsNullOrWhiteSpace(ucSupplier.UC_Code))
            {
                if (!ucSupplier.SelectData())
                {
                    DSP_MSG("E101", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                    ucSupplier.SetFocus();
                    return false;
                }
                else if (!psks0103ibl.MakerBrandSelect("MakerCD", "M_MakerBrand", "Where DataSourseMakerCD ='" + ucSupplier.UC_Code + "'"))
                {
                    DSP_MSG("E101", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                    ucSupplier.SetFocus();
                    return false;
                }
            }
            else
            {
                DSP_MSG("E102", "名前", string.Empty, string.Empty, string.Empty, string.Empty);
                ucSupplier.SetFocus();
                return false;
            }

            //ErrChkFlg = "0";
            foreach (DataGridViewRow row1 in dgvImportRireki.Rows)
            {
                DataGridViewCheckBoxCell chk1 = row1.Cells["chk1"] as DataGridViewCheckBoxCell;
                DataGridViewCheckBoxCell chk2 = row1.Cells["chk2"] as DataGridViewCheckBoxCell;
                DataGridViewCheckBoxCell chk3 = row1.Cells["chk3"] as DataGridViewCheckBoxCell;

                if (chk1.Value == null) chk1.Value = false;
                if (chk2.Value == null) chk2.Value = false;
                if (chk3.Value == null) chk3.Value = false;

                if (Convert.ToBoolean(chk1.Value) != false && Convert.ToBoolean(chk2.Value) != false)// check「結果」「エラー」
                {
                    DSP_MSG("E126", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                    return false;
                }
                else if (Convert.ToBoolean(chk3.Value) == false)//Check File Type and Folder path
                {
                    if (FilePath.Count <= 0)
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

            if (string.IsNullOrWhiteSpace(cboPatternCD.Text))
            {
                MessageBox.Show("フォーマット Doesn't Exists", "メッセージ");
                return false;
            }

            if (!(Have_opened_Excel(FilePath)))  // Added by Phyoe Gyi  in 03082019   --- for Checking the imported excel are opened or not?
            {
                MessageBox.Show("Some of the imported files are being used by another process. please close those files to be able to import");
                return false;
            }
            return true;
        }

        /// <summary>get file path
        /// 
        /// </summary>
        /// <returns></returns>
        private ArrayList getFilePath()
        {
            ArrayList arr = new ArrayList();

            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            DirectoryInfo d = new DirectoryInfo(folderPath);
            string[] Files;

            if (cboFileType.SelectedValue.Equals("1"))
            {
                Files = Directory.GetFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly).Where(s => s.EndsWith(".csv")).ToArray();
            }
            else
            {
                Files = Directory.GetFiles(folderPath, "*.*", SearchOption.TopDirectoryOnly).Where(s => s.EndsWith("ls") || s.EndsWith("sx")).ToArray();
            }

            foreach (string file in Files)
            {
                if (Path.GetFileName(file).StartsWith(FileName))
                {

                    arr.Add(file);
                }
            }

            return arr;
        }

        private void btnCSVExport_Click(object sender, EventArgs e)
        {
            if(dgvFileDetail.Rows.Count>0)
            {
                DataTable dtCSV = dtmakerinput.Copy();
                if(ImportFileName.Contains(".csv"))
                    ImportFileName = ImportFileName.Replace(".csv", "");
                else
                    ImportFileName = ImportFileName.Replace(".xlsx", "");

                string folderPath = "C:\\SMS\\";
                if (!Directory.Exists(folderPath))
                {
                    Directory.CreateDirectory(folderPath);
                }
                SaveFileDialog savedialog = new SaveFileDialog();
                savedialog.Filter = "CSV|*.csv";
                savedialog.Title = "Save";
                savedialog.FileName = ErrChkFlg=="2"? ImportFileName+"_Error": ImportFileName + "_Result";
                savedialog.InitialDirectory = folderPath;
                savedialog.RestoreDirectory = true;
                if (savedialog.ShowDialog() == DialogResult.OK)
                {
                    ChangedColumnName(dtCSV);
                    if (Path.GetExtension(savedialog.FileName).Contains("csv"))
                    {
                        CsvWriter csvwriter = new CsvWriter();
                        csvwriter.WriteCsv(dtCSV, savedialog.FileName, Encoding.GetEncoding(932));
                    }
                    Process.Start(Path.GetDirectoryName(savedialog.FileName));
                }
            }
        }
        public void ChangedColumnName(DataTable dtCSV)
        {
            dtCSV.Columns["vm_hanbai_shohin"].ColumnName = "商品名";
            dtCSV.Columns["InportSu"].ColumnName = "在庫数";
            dtCSV.Columns["HenkanNouki"].ColumnName = "納期";
            dtCSV.Columns["ErrorText"].ColumnName = "エラー内容・対応方法";

            dtCSV.Columns.Remove("SetKBN");
        }
    }
}
