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
using SMS_BL;
using System.Data;
using System.Threading;
using Sms_Prog;
using SMS.CustomControls;

namespace SMS.Inventory
{
    public partial class frmPSKS0104I : SMS_Base
    {
        Login_Info loginInfo;
        M_SiiresakiZaiko_Entity msze;
        DataTable dt_PSKS0104I;
        PSKS0104I_BL psks0104ibl;
        ItemSearch_Entity ise;
        T_MakerZaiko_Entity tmze;
        L_Log_Entity lle;
        int SearchMode = 0;// 1 - maker, 2 - brand, 3 - item, 4 - makershohin, 5 - jan, 6 - sku, 7 - jan add
        string vm_hanbai_shohin = string.Empty;

        /// <summary>
        /// Add ComboBox Data
        /// </summary>
        private class Item
        {
            public string Value;
            public string Code;
            public Item(string code, string value)
            {
                Code = code; Value = value;
            }
            public override string ToString()
            {
                // Generates the text shown in the combo box
                return Code + ":" + Value;
            }
        }

        #region constructor
        public frmPSKS0104I()
        {
            InitializeComponent();
        }

        public frmPSKS0104I(Login_Info loginInfo)
            : base(loginInfo)
        {
            InitializeComponent();
            this.loginInfo = loginInfo;
        }
        #endregion

        /// <summary>
        /// form load
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmPSKS0104I_Load(object sender, EventArgs e)
        {
            psks0104ibl = new PSKS0104I_BL();

            this.FormName = "メーカー在庫確定入力";
            lblMode.Visible = false;
            FunctionButtonDisabled(2);
            FunctionButtonDisabled(3);
            FunctionButtonDisabled(4);
            FunctionButtonDisabled(6);
            FunctionButtonDisabled(7);
            FunctionButtonDisabled(8);
            FunctionButtonDisabled(10);
            FunctionButtonDisabled(11);
            FunctionButtonDisabled(12);

            dgvPSK0104I.EnabledColumn("colKakuteiZaikoSu,colNouki");
            dgvPSK0104I.DisabledSorting();

            cboNendo.Bind();
            cboSeason.Bind();

            ucSupplier.SetFocus();
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
                case 5:
                    Clear();
                    break;
                case 9:
                    F9();
                    break;
                case 11:
                    Search();
                    break;
                case 12:
                    F12();
                    break;
            }
        }

        #region Handle for change KakuteiMode
        private void btnKakuteiMode(object sender, EventArgs e)
        {
            var btnSender = (Button)sender;

            if (btnSender.Text == "確定")
                KakuteiModeChange(true);
            else
                KakuteiModeChange(false);
        }

        /// <summary>
        /// Change Button color (Blue and Yellow)
        /// </summary>
        /// <param name="FlgKaku"></param>
        private void KakuteiModeChange(bool FlgKaku)
        {
            lblKakuteiMode.Visible = true;
            if (FlgKaku)
            {
                lblKakuteiMode.Text = "確定";
                btnKakutei.Enabled = false;
                btnMiKakutei.Enabled = true;
                lblKakuteiMode.BackColor = Color.FromArgb(0, 176, 240);
            }
            else
            {
                lblKakuteiMode.Text = "未確定";
                btnMiKakutei.Enabled = false;
                btnKakutei.Enabled = true;
                lblKakuteiMode.BackColor = Color.Yellow;
            }
        }

        #endregion

        #region Searching
        /// <summary>
        /// Searching
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDisplay_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(ucSupplier.UC_Code))
            {
                Search();
            }
            else
            {
                DSP_MSG("E102", "名前", string.Empty, string.Empty, string.Empty, string.Empty);
                ucSupplier.SetFocus();
            }
        }

        /// <summary>
        /// Display Data for GridView
        /// </summary>
        protected void Search()
        {
            if (!string.IsNullOrWhiteSpace(ucSupplier.UC_Code))
            {

                Cursor.Current = Cursors.WaitCursor;
                DataTable dtmakerinput = new DataTable();
                ise = GetData();
                //dt_PSKS0104I = new DataTable();
                dt_PSKS0104I = psks0104ibl.PSKS0104I_Select(ise);
                dgvPSK0104I.DataSource = dt_PSKS0104I;

                DisabledCtrl();


            }
            else
            {
                DSP_MSG("E102", "名前", string.Empty, string.Empty, string.Empty, string.Empty);
                ucSupplier.SetFocus();
            }

        }

        #endregion



        /// <summary>
        /// Get Searching Data
        /// </summary>
        /// <returns></returns>
        private ItemSearch_Entity GetData()
        {
            ise = new ItemSearch_Entity();
            ise.makercd = ucSupplier.UC_Code;
            if (!string.IsNullOrWhiteSpace(txtInportDate.Text.Replace("/", "")))
                ise.InportDate = txtInportDate.Text;
            ise.Zaiko0 = chkZaikoSuZero.Checked ? "1" : string.Empty;
            ise.FirstInport = chkJANCDcut.Checked ? "1" : string.Empty;
            ise.brand = ucBrand.UC_Code;
            ise.catalog = txtCatalogue.Text;
            ise.nendo = cboNendo.SelectedValue.ToString();
            ise.season = cboSeason.SelectedValue.ToString();
            ise.itemCode = ucITEM.UC_Code;
            ise.skucd = ucSKUCD.UC_Code;
            ise.makerShohin = ucMakerItem.UC_Code;
            ise.jancd = ucJANCD1.UC_Code.TrimStart(',').TrimEnd(',');
            if (chkSearch.Checked && (!string.IsNullOrWhiteSpace(ucITEM.UC_Code) || !string.IsNullOrWhiteSpace(ucSKUCD.UC_Code) || !string.IsNullOrWhiteSpace(ucJANCD1.UC_Code)))
            {
                if (rdoItem.Checked)
                    ise.itemCheck = "1";
                else
                    ise.makerCheck = "1";
            }

            return ise;
        }

        /// <summary>
        /// Error Check function
        /// </summary>
        /// <returns></returns>
        private bool ErrorCheck2()
        {
            //ktp
            foreach (DataGridViewRow dr in dgvPSK0104I.Rows)
            {
                if (string.IsNullOrWhiteSpace(dr.Cells["colKakuteiZaikoSu"].Value.ToString()))
                {
                    dr.Cells["colKakuteiZaikoSu"].Value = 0;
                }
                if (Convert.ToInt32(dr.Cells["colKakuteiZaikoSu"].Value.ToString().Replace(",", "")) < 0)
                {
                    DSP_MSG("E109", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                    dgvPSK0104I.CurrentCell = dgvPSK0104I[dr.Cells["colKakuteiZaikoSu"].ColumnIndex, dr.Index];
                    return false;
                }

            }
            if (string.IsNullOrWhiteSpace(ucSupplier.UC_Code) || string.IsNullOrWhiteSpace(ucSupplier.UC_Name))
            {
                DSP_MSG("E102", "名前", string.Empty, string.Empty, string.Empty, string.Empty);
                return false;
            }

            return true;
        }

        #region Add new record to Gridview
        private void F10()
        {
            if (GvAddRow_Check())
            {
                dt_PSKS0104I.Rows.Add();
                dt_PSKS0104I.Rows[dt_PSKS0104I.Rows.Count - 1]["No"] = dt_PSKS0104I.Rows.Count.ToString();
                dt_PSKS0104I.Rows[dt_PSKS0104I.Rows.Count - 1]["JanCD"] = ucJANCD2.UC_Code;
                dt_PSKS0104I.Rows[dt_PSKS0104I.Rows.Count - 1]["vm_hanbai_shohin"] = vm_hanbai_shohin;// JANCD nae ItemName ko get load

                dt_PSKS0104I.Rows[dt_PSKS0104I.Rows.Count - 1]["InportSu"] = "0";
                if (string.IsNullOrWhiteSpace(txtZaikoSu.Text))
                {
                    dt_PSKS0104I.Rows[dt_PSKS0104I.Rows.Count - 1]["UpdateSu"] = "0";
                }
                else
                {
                    dt_PSKS0104I.Rows[dt_PSKS0104I.Rows.Count - 1]["UpdateSu"] = txtZaikoSu.Text; // Convert.ToInt32(txtZaikoSu.Text);
                }
                dt_PSKS0104I.Rows[dt_PSKS0104I.Rows.Count - 1]["HenkanNouki"] = null;
                dt_PSKS0104I.Rows[dt_PSKS0104I.Rows.Count - 1]["dgvMsg"] = null;
                dgvPSK0104I.DataSource = dt_PSKS0104I;

                ucJANCD2.UC_Code = string.Empty;
                ucJANCD2.UC_Name = string.Empty;
                txtZaikoSu.Text = "";

            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            F10();
        }

        /// <summary>
        /// check to add new record
        /// </summary>
        /// <returns></returns>
        private bool GvAddRow_Check()
        {
            if (!string.IsNullOrWhiteSpace(ucSupplier.UC_Code))
            {
                txtZaikoSu.Text = string.IsNullOrWhiteSpace(txtZaikoSu.Text) ? "0" : txtZaikoSu.Text;
                if (Convert.ToInt32(txtZaikoSu.Text.Replace(",", "")) < 0)
                {
                    DSP_MSG("E109", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                    return false;
                }

                if (!string.IsNullOrWhiteSpace(ucJANCD2.UC_Code))
                {
                    DataTable dt = new DataTable();
                    dt = psks0104ibl.JANCD_ExistsCheck(ucJANCD2.UC_Code);    //Check Exist in Sugoraku
                    if (dt.Rows.Count <= 0)
                    {
                        DSP_MSG("E101", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                        return false;
                    }
                    else
                    {
                        vm_hanbai_shohin = dt.Rows[0]["VM_SHOHIN"].ToString();
                    }
                }
                else
                {
                    DSP_MSG("E102", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                    ucJANCD2.SetFocus();
                    return false;
                }

                tmze = new T_MakerZaiko_Entity();
                tmze.MakerCD = ucSupplier.UC_Code;
                tmze.JanCD = ucJANCD2.UC_Code;
                DataRow[] drJanDuplicate = dt_PSKS0104I.Select("JANCD='" + ucJANCD2.UC_Code + "'");
                if (psks0104ibl.T_MakerZaiko_ExistCheck(tmze).Rows.Count > 0 || drJanDuplicate.Count() > 0)   // Check Exist in shouhinkanri 
                {
                    DSP_MSG("E131", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                    return false;
                }
            }
            else
            {
                DSP_MSG("E102", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                ucSupplier.SetFocus();
                return false;
            }
            return true;
        }
        #endregion

        /// <summary>
        /// Save record from Gridview
        /// </summary>
        private void F12()
        {
            if (DSP_MSG("Q101", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty).Equals("1"))
            {
                if (!string.IsNullOrWhiteSpace(ucSupplier.UC_Code))
                {
                    string DateAndTime = System.DateTime.Now.ToString();
                    if (ErrorCheck2())
                    {
                        if (dgvPSK0104I.Rows.Count > 0)
                        {
                            tmze = GetTmze();
                            lle = GetLogData();
                            psks0104ibl.psks0104I_InsertUpdate(tmze, lle);
                            EnabledCtrl();
                            Clear();
                            dt_PSKS0104I.AcceptChanges();
                        }

                    }
                    else
                    {
                        ucSupplier.SetFocus();
                    }
                }
                else
                {
                    DSP_MSG("E102", "名前", string.Empty, string.Empty, string.Empty, string.Empty);
                    ucSupplier.SetFocus();
                }
            }
        }

        /// <summary>
        /// Add from Datatable Data to Entity
        /// </summary>
        /// <returns></returns>
        private T_MakerZaiko_Entity GetTmze()
        {
            tmze = new T_MakerZaiko_Entity();
            tmze.MakerCD = ucSupplier.UC_Code;
            tmze.InsertOperator = loginInfo.OperatorCode;
            tmze.InsertDateTime = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
            tmze.KakuTeiFlg = lblKakuteiMode.Text == "確定" ? "1" : "0";
            tmze.dt1 = dt_PSKS0104I;

            return tmze;
        }

        /// <summary>
        /// Add log data to Entity
        /// </summary>
        /// <returns></returns>
        private L_Log_Entity GetLogData()
        {
            lle = new L_Log_Entity();
            lle.InsertOperator = loginInfo.OperatorCode;
            lle.Program = "PSKS0104I";
            lle.PC = loginInfo.PcName;
            lle.OperateMode = loginInfo.StartMode;
            lle.KeyItem = ucSupplier.UC_Code + " " + lblKakuteiMode.Text;

            return lle;
        }

        private void DisabledCtrl()
        {
            ucSupplier.Disabled();
            //ucBrand.Disabled();
            //ucITEM.Disabled();
            //ucMakerItem.Disabled();
            //ucJANCD1.Disabled();

            //chkSearch.Enabled = false;
            //rdoItem.Enabled = false;
            //rdoMakerItemCD.Enabled = false;
            //chkJANCDcut.Enabled = false;
            //chkZaikoSuZero.Enabled = false;

            //txtCatalogue.Enabled = false;
            //txtInportDate.Enabled = false;
            //ucSKUCD.Disabled();

            //cboNendo.Enabled = false;
            //cboSeason.Enabled = false;

            //btnDisplay.Enabled = false;
            FunctionButtonEnabled(12);
        }

        private void EnabledCtrl()
        {
            ucSupplier.Enable();
            //ucBrand.Enable();
            //ucITEM.Enable();
            //ucMakerItem.Enable();
            //ucJANCD1.Enable();

            //chkSearch.Enabled = true;
            //rdoItem.Enabled = true;
            //rdoMakerItemCD.Enabled = true;
            //chkJANCDcut.Enabled = true;
            //chkZaikoSuZero.Enabled = true;

            //txtCatalogue.Enabled = true;
            //txtInportDate.Enabled = true;
            //ucSKUCD.Enable();

            //cboNendo.Enabled = true;
            //cboSeason.Enabled = true;

            //btnDisplay.Enabled = true;
            FunctionButtonDisabled(12);
        }


        /// <summary>
        /// Clear Data
        /// </summary>
        private void Clear()
        {
            ucSupplier.UC_Code = string.Empty;
            ucSupplier.UC_Name = string.Empty;

            ucBrand.UC_Code = string.Empty;
            ucBrand.UC_Name = string.Empty;

            ucITEM.UC_Code = string.Empty;
            ucITEM.UC_Name = string.Empty;

            ucMakerItem.UC_Code = string.Empty;
            ucMakerItem.UC_Name = string.Empty;

            txtInportDate.Text = string.Empty;
            txtCatalogue.Text = string.Empty;

            ucJANCD1.UC_Code = string.Empty;
            ucJANCD1.UC_Name = string.Empty;

            ucSKUCD.UC_Code = string.Empty;
            ucSKUCD.UC_Name = string.Empty;

            ucJANCD2.UC_Code = string.Empty;
            ucJANCD2.UC_Name = string.Empty;

            cboNendo.SelectedValue = "-1";
            cboSeason.SelectedValue = "-1";

            txtZaikoSu.Text = string.Empty;
            dgvPSK0104I.DataSource = null;

            chkJANCDcut.Checked = false;
            chkZaikoSuZero.Checked = false;
            chkSearch.Checked = false;

            rdoItem.Checked = true;
            rdoMakerItemCD.Checked = false;

            EnabledCtrl();

            //dt_PSKS0104I.Clear();
            ucSupplier.SetFocus();



            lblKakuteiMode.Text = string.Empty;
            lblKakuteiMode.BackColor = Color.FromArgb(0, 176, 240);
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
            else if (ucsearch.UC_Type == UC_Search.Type.ITEM)
                SearchMode = 3;
            else if (ucsearch.UC_Type == UC_Search.Type.MakerShohinCD)
                SearchMode = 4;
            else if (ucsearch.UC_Type == UC_Search.Type.MultiJANCD)
                SearchMode = 5;
            else if (ucsearch.UC_Type == UC_Search.Type.SKUCD)
                SearchMode = 6;
            else if (ucsearch.UC_Type == UC_Search.Type.JANCD)
                SearchMode = 7;

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
                ucITEM.ShowSearchForm();
            else if (SearchMode == 4)
                ucMakerItem.ShowSearchForm();
            else if (SearchMode == 5)
                ucJANCD1.ShowSearchForm();
            else if (SearchMode == 6)
                ucSKUCD.ShowSearchForm();
            else if (SearchMode == 7)
                ucJANCD2.ShowSearchForm();
        }

        #region keydown event
        private void frmPSKS0104I_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F11)
                Search();
            else if (e.KeyCode == Keys.F10)
                F10();
        }

        private void ucSupplier_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (!string.IsNullOrWhiteSpace(ucSupplier.UC_Code))
                {
                    if (ucSupplier.SelectData())
                    {
                        msze = new M_SiiresakiZaiko_Entity();
                        msze.SiiresakiCD = ucSupplier.UC_Code;
                        DataTable dt = psks0104ibl.M_SiiresakiZaiko_Select(msze);
                        if (dt.Rows.Count > 0)
                        {
                            if (dt.Rows[0]["TorikomiFixedFLG"].ToString() == "1")
                            {
                                KakuteiModeChange(true);
                            }
                            else
                            {
                                KakuteiModeChange(false);
                            }
                        }
                    }
                    else
                    {
                        DSP_MSG("E101", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                        lblKakuteiMode.Text = "";
                        lblKakuteiMode.BackColor = Color.FromArgb(0, 176, 240);
                        btnKakutei.Enabled = true;
                        btnMiKakutei.Enabled = true;
                    }
                }
                else
                {
                    DSP_MSG("E102", "名前", string.Empty, string.Empty, string.Empty, string.Empty);
                    ucSupplier.SetFocus();
                }
            }
        }

        private void ucJANCD2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                DataTable dt = psks0104ibl.JANCD_ExistsCheck(ucJANCD2.UC_Code);
                if (dt.Rows.Count > 0)
                {
                    ucJANCD2.UC_Name = dt.Rows[0]["VM_SHOHIN"].ToString();
                }
                else
                {
                    DSP_MSG("E101", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                    return;
                }
            }
        }
        #endregion

        private void dgvPSK0104I_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == dgvPSK0104I.Columns["colNouki"].Index)
                dgvPSK0104I.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            else if (e.ColumnIndex == dgvPSK0104I.Columns["colKakuteiZaikoSu"].Index)
                dgvPSK0104I.ImeMode = System.Windows.Forms.ImeMode.Disable;
        }

        /// <summary>
        /// Change Back Cell Color to blue
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvPSK0104I_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (dgvPSK0104I.Rows[e.RowIndex].Cells["coluchiwake"].Value.ToString() == "2")
            {
                dgvPSK0104I.Rows[e.RowIndex].Cells["colNo"].Style.ForeColor = System.Drawing.Color.Blue;
                dgvPSK0104I.Rows[e.RowIndex].Cells["colJANCD"].Style.ForeColor = System.Drawing.Color.Blue;
                dgvPSK0104I.Rows[e.RowIndex].Cells["colItemName"].Style.ForeColor = System.Drawing.Color.Blue;
                dgvPSK0104I.Rows[e.RowIndex].Cells["colToriKomiDateTime"].Style.ForeColor = System.Drawing.Color.Blue;
                dgvPSK0104I.Rows[e.RowIndex].Cells["colToriKomiNo"].Style.ForeColor = System.Drawing.Color.Blue;
                dgvPSK0104I.Rows[e.RowIndex].Cells["colTokukiJikou"].Style.ForeColor = System.Drawing.Color.Blue;
            }
        }


    }
}
