using SMS_BL;
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

namespace SMS.Inventory
{
    public partial class frmPSKS0101M : SMS_Base
    {
        DataTable dtKaMoku = new DataTable();
        DataTable dtNouki = new DataTable();
        DataTable dtSuuryo = new DataTable();
        Login_Info loginInfo;
        M_MakerZaiko_H_Entity makerH_entity;
        MakerZaiko_D_Entity makerD_entity;
        M_NoukiHenkan_Entity nouki_entity;
        SuuryoHenKan_Entity suuryo_entity;
        L_Log_Entity LLogEntity;
        PSKS0101M_BL sks0101M_bl = new PSKS0101M_BL();
        int SearchMode = 1;//1 - Pattern, 2 - Copy Pattern, 3 - Others

        #region constructor
        /// <summary>
        /// default
        /// </summary>
        public frmPSKS0101M()
        {
            InitializeComponent();
        }

        public frmPSKS0101M(Login_Info loginInfo)
            : base(loginInfo)
        {
            InitializeComponent();
            this.loginInfo = loginInfo;
            //SetDesignerFunction();
            NewMode();
        }

        #endregion

        /// <summary>
        /// form load
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmPSKS0101M_Load(object sender, EventArgs e)
        {
            FormName = "メーカー在庫フォーマット設定マスタ";
            functionClick(2);

            dgvKamoku.DisabledColumn("colNo");
            dgvdelivery.DisabledColumn("colNum");
            dgvSuuryoHenKan.DisabledColumn("colSEQ");

            FunctionButtonDisabled(7);
            FunctionButtonDisabled(8);
            FunctionButtonDisabled(10);
            FunctionButtonDisabled(11);
        }

        #region Initialize 300 rows to gridviews
        public void CreateDataTableforKouMokuSetting()
        {
            dtKaMoku = new DataTable();

            dtKaMoku.Columns.Add("SEQ");
            dtKaMoku.Columns.Add("ItemName");
            dtKaMoku.Columns.Add("ItemProperty");
            for (int i = 0; i < 300; i++)
            {
                dtKaMoku.Rows.Add();
                dtKaMoku.Rows[i]["SEQ"] = i + 1;
                dtKaMoku.Rows[i]["ItemProperty"] = "0";
            }
            dgvKamoku.DataSource = dtKaMoku;
        }

        public void CreateDataTableforNoukiSetting()
        {
            dtNouki = new DataTable();

            dtNouki.Columns.Add("SEQ");
            dtNouki.Columns.Add("Maker");
            dtNouki.Columns.Add("Conversion");

            for (int i = 0; i < 300; i++)
            {
                dtNouki.Rows.Add();
                dtNouki.Rows[i]["SEQ"] = i + 1;
            }
            dgvdelivery.DataSource = dtNouki;
        }

        public void CreateDataTableforSuuryoHenKan()
        {
            dtSuuryo = new DataTable();

            dtSuuryo.Columns.Add("SEQ");
            dtSuuryo.Columns.Add("MakerSuuryo");
            dtSuuryo.Columns.Add("ConversionSuuryo");

            for (int i = 0; i < 300; i++)
            {
                dtSuuryo.Rows.Add();
                dtSuuryo.Rows[i]["SEQ"] = i + 1;
            }
            dgvSuuryoHenKan.DataSource = dtSuuryo;
        }

        #endregion

        /// <summary>
        /// Bind combo of grid
        /// </summary>
        private void BindCombo()
        {

            DataTable dt = new DataTable();
            dt.Columns.Add("ID");
            dt.Columns.Add("PropertyName");

            dt.Rows.Add("0", string.Empty);
            dt.Rows.Add("1", "01:JANCD");
            dt.Rows.Add("2", "02:在庫数(加算取込)");
            dt.Rows.Add("3", "03:在庫数(減算取込)");
            dt.Rows.Add("4", "04:在庫数(無視)");

            dt.Rows.Add("5", "05:メーカー商品CD(無視)");
            dt.Rows.Add("6", "06:カラー(無視)");
            dt.Rows.Add("7", "07:サイズ(無視)");
            dt.Rows.Add("8", "08:ブランド名(無視)");

            dt.Rows.Add("9", "09:納期(変換取込)");
            dt.Rows.Add("10", "10:納期(無変換取込)");
            dt.Rows.Add("11", "11:納期(無視)");
            dt.Rows.Add("12", "12:備考(無視)");
            dt.Rows.Add("13", "13:その他・空白(無視)");

            DataGridViewComboBoxColumn col = (DataGridViewComboBoxColumn)dgvKamoku.Columns["colZokuSei"];
            col.DataPropertyName = "ItemProperty";
            col.DisplayMember = "PropertyName";
            col.ValueMember = "ID";

            ((DataGridViewComboBoxColumn)dgvKamoku.Columns["colZokuSei"]).DataSource = dt;
        }

        //private void SetDesignerFunction()
        //{
        //    dgvKamoku.DisabledColumn("colNo");
        //    dgvdelivery.DisabledColumn("colNum");
        //    dgvSuuryoHenKan.DisabledColumn("colSEQ");

        //    FunctionButtonDisabled(7);
        //    FunctionButtonDisabled(8);
        //    FunctionButtonDisabled(10);
        //    FunctionButtonDisabled(11);
        //}

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
                case 2:
                    Clear();
                    ucPattern.Enable();
                    txtPatternName.Enable();
                    ucPattern.UC_SearchEnable = false;
                    ucPattern.SetFocus();
                    ucCopyPattern.Enable();
                    dgvKamoku.ReadOnly = false;
                    dgvdelivery.ReadOnly = false;
                    dgvSuuryoHenKan.ReadOnly = false;
                    dgvKamoku.Columns[0].ReadOnly = true;
                    dgvdelivery.Columns[0].ReadOnly = true;
                    dgvSuuryoHenKan.Columns[0].ReadOnly = true;
                    break;
                case 3:
                case 11:
                    Clear();
                    txtPatternName.Enabled = true;
                    ucPattern.Enable();
                    ucPattern.SetFocus();
                    ucCopyPattern.Disabled();
                    dgvKamoku.ReadOnly = true;
                    dgvdelivery.ReadOnly = true;
                    dgvSuuryoHenKan.ReadOnly = true;
                    break;
                case 4:
                case 5:
                    Clear();
                    dgvKamoku.ReadOnly = true;
                    dgvdelivery.ReadOnly = true;
                    dgvSuuryoHenKan.ReadOnly = true;
                    txtPatternName.Enabled = false;
                    ucPattern.Enable();
                    ucPattern.SetFocus();
                    ucCopyPattern.Disabled();
                    break;
                case 6:
                    Clear();
                    ucPattern.SetFocus();
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
        /// Clean data
        /// </summary>
        private void Clear()
        {
            CreateDataTableforKouMokuSetting();
            CreateDataTableforNoukiSetting();
            CreateDataTableforSuuryoHenKan();
            BindCombo();
            ucPattern.UC_Code = string.Empty;
            ucCopyPattern.UC_Code = string.Empty;
            ucCopyPattern.UC_Name = string.Empty;
            txtPatternName.Text = string.Empty;
            ucCopyPattern.Enable();
            ucPattern.Enable();
            txtPatternName.Enable();
            ucPattern.UC_SearchEnable = false;
            if (FormMode == "照会" || FormMode == "削除")
            {
                dgvKamoku.ReadOnly = true;
                dgvdelivery.ReadOnly = true;
                dgvSuuryoHenKan.ReadOnly = true;
                txtPatternName.Enabled = false;
                ucPattern.Enable();
                ucPattern.SetFocus();
                ucCopyPattern.Disabled();

            }
            else if (FormMode == "修正")
            {
                txtPatternName.Enabled = true;
                ucPattern.Enable();
                ucPattern.SetFocus();
                ucCopyPattern.Disabled();
                dgvKamoku.ReadOnly = true;
                dgvdelivery.ReadOnly = true;
                dgvSuuryoHenKan.ReadOnly = true;
            }
            
        }

        #region Save adding or editing records
        private void F12()
        {
            if (ErrorCheck())
            {
                string confirmMessageID = FormMode.Equals("削除") ? "Q102" : "Q101";
                string msg = DSP_MSG(confirmMessageID, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);

                if (msg.Equals("1"))
                {
                    if (ExecutionProcess())
                    {
                        Clear();
                    }
                    else
                    {
                        DSP_MSG("E015", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                    }
                }
            }
         }

        private bool ExecutionProcess()
        {
            makerH_entity = GetMakerH_Entity();
            makerD_entity = GetMakerD_Entity();
            nouki_entity = GetNouki_Entity();
            suuryo_entity = GetSuuryo_Entity();          
            LLogEntity = new L_Log_Entity();
            LLogEntity.InsertOperator = loginInfo.OperatorCode;
            string dateandTime = DateTime.Now.ToString();
            LLogEntity.OperateDate = makerH_entity.InsertDateTime = nouki_entity.InsertDateTime = suuryo_entity.InsertDateTime = dateandTime;
            LLogEntity.Program = "PSKS0101M";
            LLogEntity.PC = loginInfo.PcName;
            LLogEntity.OperateMode = FormMode;
            LLogEntity.KeyItem = ucPattern.UC_Code;
            makerH_entity.InsertOperator = loginInfo.OperatorCode;
            nouki_entity.InsertOperator = loginInfo.OperatorCode;
            suuryo_entity.InsertOperator = loginInfo.OperatorCode;
            bool result = false;
            switch (FormMode)
            {
                case "新規":
                    result = sks0101M_bl.MakerZaiko_Insert(makerH_entity, makerD_entity, nouki_entity, LLogEntity, suuryo_entity);
                    break;
                case "修正":
                    makerH_entity.UpdateOperator = loginInfo.OperatorCode;
                    nouki_entity.UpdateOperator = loginInfo.OperatorCode;
                    suuryo_entity.UpdateOperator = loginInfo.OperatorCode;
                    makerH_entity.UpdateDateTime = nouki_entity.UpdateDateTime =suuryo_entity.UpdateDateTime = dateandTime;
                    result = sks0101M_bl.MakerZaiko_Update(makerH_entity, makerD_entity, nouki_entity, LLogEntity, suuryo_entity);
                    break;
                case "削除":
                    result = sks0101M_bl.MakerZaiko_Delete(makerH_entity, makerD_entity, nouki_entity, LLogEntity, suuryo_entity);
                    break;
            }
            return result;
           
        }

        /// <summary>
        /// check error
        /// </summary>
        /// <returns></returns>
        private bool ErrorCheck()
        {
            if (string.IsNullOrWhiteSpace(ucPattern.UC_Code))//|| string.IsNullOrWhiteSpace(txtPatternName .Text ))
            {
                DSP_MSG("E102", "名前", string.Empty, string.Empty, string.Empty, string.Empty);
                ucPattern.SetFocus();
                return false;
            }
            else
            {
                if (ucPattern.SelectData())
                {
                    if (FormMode == "新規")
                    {
                        DSP_MSG("E107", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                        ucPattern.Enable();
                        ucPattern.SetFocus();
                        return false;
                    }
                }
            }

            if (string.IsNullOrWhiteSpace(txtPatternName.Text))
            {
                DSP_MSG("E102", "名前", string.Empty, string.Empty, string.Empty, string.Empty);
                txtPatternName.Focus();
                return false;
            }

            //***BugNo5***
            //No Need to check anymore,already check in combobox selectedIndexChanged
            //DataRow[] dr1 = dtKaMoku.Select("ItemProperty = '1' ");
            //if (dr1.Count() > 1)
            //{
            //    DSP_MSG("E124", "名前", string.Empty, string.Empty, string.Empty, string.Empty);
            //    return false;
            //}
            //end

            //DataRow[] drCheck1 = dtKaMoku.Select("ItemProperty = '9'");
            //DataRow[] drCheck2 = dtKaMoku.Select("ItemProperty = '10'");
            //if (drCheck1.Count() > 0 && drCheck2.Count() > 0)
            //{
            //    DSP_MSG("E129", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
            //    dgvKamoku.Select();
            //    dgvKamoku.CurrentCell = dgvKamoku[dgvKamoku.Columns["colZokuSei"].Index, Convert.ToInt16(drCheck1[0]["SEQ"].ToString()) - 1];
            //    dgvdelivery.ClearSelection();
            //    return false;
            //}

            DataRow[] dr3 = dtKaMoku.Select("(ItemProperty = '' OR ItemProperty IS NULL) AND ItemProperty = '' OR ItemName IS NULL ");

            if (dr3.Count() == 300)
            {
                DSP_MSG("E125", "メーカー在庫フォーマットDetail", string.Empty, string.Empty, string.Empty, string.Empty);
                dgvKamoku.CurrentCell = dgvKamoku[dgvKamoku.Columns["colName"].Index, Convert.ToInt16(dr3[0]["SEQ"].ToString()) - 1];
                return false;
            }


            //***BugNo26***
            //Set Cursor to duplicate cell,change checking logic

            //DataTable distinct = dtNouki.DefaultView.ToTable(true, "Maker");
            //DataRow[] dr = dtNouki.Select("Maker <> '' OR Maker IS NOT NULL");
            //if (distinct.Rows.Count != (dr.Count() + 1))
            //{
            //    DSP_MSG("E105", "メーカー値", string.Empty, string.Empty, string.Empty, string.Empty);
            //    return false;
            //}

            DataRow[] drs = dtNouki.Select("Maker <> '' OR Maker IS NOT NULL");
            if (drs.Count() > 0)
            {
                DataTable dt1 = drs.CopyToDataTable();
                foreach (DataRow dr in dt1.Rows)
                {
                    drs = dt1.Select("Maker = '" + dr["Maker"] + "'");
                    if (drs.Count() > 1)
                    {
                        DSP_MSG("E105", "メーカー値", string.Empty, string.Empty, string.Empty, string.Empty);
                        dgvdelivery.Select();
                        dgvdelivery.CurrentCell = dgvdelivery[dgvdelivery.Columns["colNouki"].Index, Convert.ToInt16(drs[0]["SEQ"].ToString()) - 1];//move cursor to duplicate cell
                        return false;
                    }
                }
            }
            //end

            DataTable suuryo = dtSuuryo.DefaultView.ToTable(true, "MakerSuuryo");
            DataRow[] dt = dtSuuryo.Select("MakerSuuryo <> '' OR MakerSuuryo IS NOT NULL");
            if (suuryo.Rows.Count != (dt.Count() + 1))
            {
                DSP_MSG("E105", "メーカー値", string.Empty, string.Empty, string.Empty, string.Empty);
                dgvSuuryoHenKan.Select();
                dgvSuuryoHenKan.CurrentCell = dgvSuuryoHenKan[dgvSuuryoHenKan.Columns["colMakerSuuryo"].Index, Convert.ToInt16(drs[0]["SEQ"].ToString()) - 1];
                return false;
            }


            foreach (DataRow rows in dtNouki.Rows)
            {
                if (string.IsNullOrWhiteSpace(rows["Maker"].ToString()))
                {
                    rows["Conversion"] = "";
                }
            }
            foreach (DataRow rows in dtKaMoku.Rows)
            {
                if (rows["ItemProperty"].ToString() == string.Empty)
                {
                    rows["ItemProperty"] = "0";
                }
                if (rows["ItemProperty"].ToString() != "0" && string.IsNullOrWhiteSpace(rows["ItemName"].ToString()))
                {
                    DSP_MSG("E102", "名前", string.Empty, string.Empty, string.Empty, string.Empty);
                    dgvKamoku.CurrentCell = dgvKamoku[dgvKamoku.Columns["colName"].Index, Convert.ToInt16(rows["SEQ"].ToString()) - 1];
                    return false;
                }
            }
            foreach (DataRow rows in dtNouki.Rows)
            {
                if (!string.IsNullOrWhiteSpace(rows["Maker"].ToString()) && string.IsNullOrWhiteSpace(rows["Conversion"].ToString()))
                {
                    DSP_MSG("E102", "納期変換", string.Empty, string.Empty, string.Empty, string.Empty);
                    dgvdelivery.CurrentCell = dgvdelivery[dgvdelivery.Columns["colHenkanNouki"].Index, Convert.ToInt16(rows["SEQ"].ToString()) - 1];
                    return false;
                }
            }

            foreach (DataRow rows in dtSuuryo.Rows)
            {
                if (!string.IsNullOrWhiteSpace(rows["MakerSuuryo"].ToString()) && string.IsNullOrWhiteSpace(rows["ConversionSuuryo"].ToString()))
                {
                    DSP_MSG("E102", "納期変換", string.Empty, string.Empty, string.Empty, string.Empty);
                    dgvSuuryoHenKan.CurrentCell = dgvSuuryoHenKan[dgvSuuryoHenKan.Columns["colConversionSuuryo"].Index, Convert.ToInt16(rows["SEQ"].ToString()) - 1];
                    return false;
                }
            }

            foreach (DataRow rows in dtKaMoku.Rows)
            {
                if ((rows["ItemProperty"].ToString() == "0" || string.IsNullOrWhiteSpace(rows["ItemProperty"].ToString())) && !string.IsNullOrWhiteSpace(rows["ItemName"].ToString()))
                {
                    DSP_MSG("E102", "名前", string.Empty, string.Empty, string.Empty, string.Empty);
                    dgvKamoku.Select();
                    dgvKamoku.CurrentCell = dgvKamoku[dgvKamoku.Columns["colZokuSei"].Index, Convert.ToInt16(rows["SEQ"].ToString()) - 1];
                    dgvdelivery.ClearSelection();
                    return false;
                }
            }


            DataRow[] dtJanCDCheck = dtKaMoku.Select("ItemProperty=1");
            if (dtJanCDCheck.Count() < 1)
            {
                DSP_MSG("E102", "名前", string.Empty, string.Empty, string.Empty, string.Empty);
                return false;
            }
            return true;
        }

        #endregion

        #region Get Searching Data
        private MakerZaiko_D_Entity GetMakerD_Entity()
        {
            makerD_entity = new MakerZaiko_D_Entity();
            makerD_entity.PatternCD = ucPattern.UC_Code;
            makerD_entity.dtKamoKu = dtKaMoku;

            return makerD_entity;
        }

        private M_NoukiHenkan_Entity GetNouki_Entity()
        {
            nouki_entity = new M_NoukiHenkan_Entity();
            nouki_entity.PatternCD = ucPattern.UC_Code;
            nouki_entity.dtNouki = dtNouki;

            return nouki_entity;
        }

        private M_MakerZaiko_H_Entity GetMakerH_Entity()
        {
            makerH_entity = new M_MakerZaiko_H_Entity();
            makerH_entity.KAICD = loginInfo.CompanyCode;
            makerH_entity.PatternCD = ucPattern.UC_Code; ;
            makerH_entity.PatternName = txtPatternName.Text;

            return makerH_entity;
        }

        private SuuryoHenKan_Entity GetSuuryo_Entity()
        {
            suuryo_entity = new SuuryoHenKan_Entity();
            suuryo_entity.PatternCD = ucPattern.UC_Code;

            for (int rowIndex = 0; rowIndex < dtSuuryo.Rows.Count; rowIndex++)
            {
                dtSuuryo.Rows[rowIndex]["ConversionSuuryo"] = dtSuuryo.Rows[rowIndex]["ConversionSuuryo"].ToString().Replace(",", "");
            }
            suuryo_entity.dtSuuryo = dtSuuryo;
            return suuryo_entity;
        }

        #endregion

        private void dgvdelivery_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            e.Control.Leave -= new EventHandler(Maker_Leave);
            if (dgvdelivery.CurrentCell.ColumnIndex == dgvdelivery.Columns["colNouki"].Index)
            {
                e.Control.Leave += new EventHandler(Maker_Leave);
            }
        }

        private void Maker_Leave(object sender, EventArgs e)
        {
            TextBox txt = sender as TextBox;
            int row = dgvdelivery.CurrentCell.RowIndex;
            if (string.IsNullOrWhiteSpace(txt.Text))
            {
                dgvdelivery.Rows[row].Cells["colHenKanNouki"].Value = string.Empty;
                return;
            }
        }


        private void dgvSuuryoHenKan_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            e.Control.Leave -= new EventHandler(SuuryoMaker_Leave);
            if (dgvSuuryoHenKan.CurrentCell.ColumnIndex == dgvSuuryoHenKan.Columns["colMakerSuuryo"].Index)
            {
                e.Control.Leave += new EventHandler(SuuryoMaker_Leave);
            }
        }

        private void SuuryoMaker_Leave(object sender, EventArgs e)
        {
            TextBox txt = sender as TextBox;
            int row = dgvSuuryoHenKan.CurrentCell.RowIndex;
            if (string.IsNullOrWhiteSpace(txt.Text))
            {
                dgvSuuryoHenKan.Rows[row].Cells["colConversionSuuryo"].Value = string.Empty;
                return;
            }
        }
       

        private void SetMaker_Data(string comCode)
        
        {
            MakerZaiko_D_Entity maker_D_data = new MakerZaiko_D_Entity();
            M_NoukiHenkan_Entity nouki_data = new M_NoukiHenkan_Entity();
            M_MakerZaiko_H_Entity maker_H_data = new M_MakerZaiko_H_Entity();
            SuuryoHenKan_Entity suuryo_data = new SuuryoHenKan_Entity();
            PSKS0101M_BL sks0101m_bl = new PSKS0101M_BL();
            DataTable dtmakerD = new DataTable();
            DataTable dtnoukidata = new DataTable();
            DataTable dtsuuryodata = new DataTable();

            if (!FormMode.Equals("新規"))
                txtPatternName.Text = ucPattern.UC_Name;

            maker_D_data.PatternCD = comCode;
            nouki_data.PatternCD = comCode;
            maker_H_data.PatternCD = comCode;
            suuryo_data.PatternCD = comCode;

            dtmakerD = sks0101m_bl.MakerZaiko_Select(maker_D_data);
            dtnoukidata = sks0101m_bl.Nouki_Select(nouki_data);
            dtsuuryodata = sks0101m_bl.Suuryo_Select(suuryo_data);

            for (int i = 1; i <= 300; i++)
            {
                DataRow[] drs = dtnoukidata.Select("SEQ = '" + i + "'");
                if (drs.Count() <= 0)
                {
                    DataRow dr = dtnoukidata.NewRow();
                    dr["SEQ"] = i.ToString();

                    dtnoukidata.Rows.Add(dr);

                    DataView dv = dtnoukidata.DefaultView;
                    dv.Sort = "SEQ ASC";
                    dtNouki = dv.ToTable();
                }
            }
            dtNouki = dtnoukidata;
            dgvdelivery.DataSource = dtNouki;


            for (int i = 1; i <= 300; i++)
            {
                DataRow[] drs = dtmakerD.Select("SEQ = '" + i + "'");
                if (drs.Count() <= 0)
                {
                    DataRow dr = dtmakerD.NewRow();
                    dr["SEQ"] = i.ToString();

                    dtmakerD.Rows.Add(dr);

                    DataView dv = dtmakerD.DefaultView;
                    dv.Sort = "SEQ ASC";
                    dtmakerD = dv.ToTable();
                }
            }
            dtKaMoku = dtmakerD;
            dgvKamoku.DataSource = dtKaMoku;

            for (int i = 1; i <= 300; i++)
            {
                DataRow[] drs = dtsuuryodata.Select("SEQ= '" + i + "'");
                if (drs.Count() <= 0)
                {
                    DataRow dr = dtsuuryodata.NewRow();
                    dr["SEQ"] = i.ToString();
                    dtsuuryodata.Rows.Add(dr);
                    DataView dv = dtsuuryodata.DefaultView;
                    dv.Sort = "SEQ ASC ";
                    dtsuuryodata = dv.ToTable();
                }
            }
            dtSuuryo = dtsuuryodata;
            dgvSuuryoHenKan.DataSource = dtsuuryodata;

            ucPattern.SetFocus();
        }



        #region  Control grid cell     ***BugNo5***

        private void dgvKamoku_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (dgvKamoku.CurrentCell.ColumnIndex == dgvKamoku.Columns["colZokuSei"].Index && e.Control is ComboBox)
            {
                ComboBox comboBox = e.Control as ComboBox;
                comboBox.SelectionChangeCommitted -= LastColumnComboSelectionChanged;
                comboBox.SelectionChangeCommitted += new EventHandler(LastColumnComboSelectionChanged);
            }
        }

        private void LastColumnComboSelectionChanged(object sender, EventArgs e)
        {
            ComboBox cbo = sender as ComboBox;
            string value1 = cbo.SelectedValue.ToString();
            if (value1.Equals("1"))//if selected value is Jan
            {
                DataRow[] dr = dtKaMoku.Select("ItemProperty = 1"); // select Jan is already exists or not
                if (dr.Count() > 0)
                {
                    if (dr[0]["SEQ"].ToString() != (dgvKamoku.CurrentCell.RowIndex + 1).ToString())// skip if combobox selected value 1:jan to 1:jan
                    {
                        DSP_MSG("E124", "名前", string.Empty, string.Empty, string.Empty, string.Empty);
                        cbo.SelectedValue = 0;//if duplicate exists,set to blank
                        dgvKamoku.CurrentCell = dgvKamoku[dgvKamoku.Columns["colZokuSei"].Index, Convert.ToInt16(dr[0]["SEQ"].ToString()) - 1];//move cursor to duplicate cell
                    }
                }
            }
            else if (value1.Equals("9"))
            {
                DataRow[] dr = dtKaMoku.Select("ItemProperty = '9'");
                DataRow[] dr1 = dtKaMoku.Select("ItemProperty = '10'");
                if (dr.Count() > 0)
                {
                    if (dr[0]["SEQ"].ToString() != (dgvKamoku.CurrentCell.RowIndex + 1).ToString())
                    {
                        DSP_MSG("E129", "名前", string.Empty, string.Empty, string.Empty, string.Empty);
                        cbo.SelectedValue = 0;
                        dgvKamoku.CurrentCell = dgvKamoku[dgvKamoku.Columns["colZokuSei"].Index, Convert.ToInt16(dr[0]["SEQ"].ToString()) - 1];
                    }
                }
                if (dr1.Count() > 0)
                {
                    if (dr1[0]["SEQ"].ToString() != (dgvKamoku.CurrentCell.RowIndex + 1).ToString())
                    {
                        DSP_MSG("E129", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                        cbo.SelectedValue = 0;
                        dgvKamoku.CurrentCell = dgvKamoku[dgvKamoku.Columns["colZokuSei"].Index, Convert.ToInt16(dr1[0]["SEQ"].ToString()) - 1];
                    }
                }
            }
            else if (value1.Equals("10"))
            {
                DataRow[] dr = dtKaMoku.Select("ItemProperty = '10'");
                DataRow[] dr1 = dtKaMoku.Select("ItemProperty = '9'");
                if (dr.Count() > 0)
                {
                    if (dr[0]["SEQ"].ToString() != (dgvKamoku.CurrentCell.RowIndex + 1).ToString())
                    {
                        DSP_MSG("E129", "名前", string.Empty, string.Empty, string.Empty, string.Empty);
                        cbo.SelectedValue = 0;
                        dgvKamoku.CurrentCell = dgvKamoku[dgvKamoku.Columns["colZokuSei"].Index, Convert.ToInt16(dr[0]["SEQ"].ToString()) - 1];
                    }
                }
                if (dr1.Count() > 0)
                {
                    if (dr1[0]["SEQ"].ToString() != (dgvKamoku.CurrentCell.RowIndex + 1).ToString())
                    {
                        DSP_MSG("E129", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                        cbo.SelectedValue = 0;
                        dgvKamoku.CurrentCell = dgvKamoku[dgvKamoku.Columns["colZokuSei"].Index, Convert.ToInt16(dr1[0]["SEQ"].ToString()) - 1];
                    }
                }
            }
        }

        #endregion

       
        /// <summary>
        /// Enter Grid cell 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgv_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == dgvKamoku.Columns["colName"].Index)
                dgvKamoku.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            else if ((e.ColumnIndex == dgvdelivery.Columns["colNouki"].Index) || (e.ColumnIndex == dgvdelivery.Columns["colHenkanNouki"].Index))
                dgvdelivery.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            else if (e.ColumnIndex == dgvSuuryoHenKan.Columns["colMakerSuuryo"].Index)
                dgvSuuryoHenKan.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            else if (e.ColumnIndex == dgvSuuryoHenKan.Columns["colConversionSuuryo"].Index)
                dgvSuuryoHenKan.ImeMode = System.Windows.Forms.ImeMode.Disable;
        }

        #region  Handle for F9 Function
        /// <summary>
        ///  handle all usercontrol to show search popup
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ucPattern_Enter(object sender, EventArgs e)
        {
            SearchMode = 1;
            if (ucPattern.UC_SearchEnable)
            {
                FunctionButtonEnabled(9);
            }
        }

        /// <summary>
        ///  handle all usercontrol to show search popup
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ucCopyPattern_Enter(object sender, EventArgs e)
        {
            SearchMode = 2;
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
                SearchMode = 3;
            }
        }

        #endregion

        #region keydown
        private void ucCopyPattern_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (!string.IsNullOrWhiteSpace(ucCopyPattern.UC_Code)) //if (e.KeyCode == Keys.Enter && ucCopyPattern.SelectData())
                {
                    if (ucCopyPattern.SelectData())
                    {
                        SetMaker_Data(ucCopyPattern.UC_Code);
                        SendKeys.Send("+{TAB}");
                        ucCopyPattern.Disabled();

                    }
                    else
                    {
                        DSP_MSG("E101", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                        CreateDataTableforKouMokuSetting();
                        CreateDataTableforNoukiSetting();
                        CreateDataTableforSuuryoHenKan();
                        BindCombo();
                        txtPatternName.Text = string.Empty;
                        ucCopyPattern.SetFocus();
                    }
                }

            }
        }

        private void txtPatternName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (string.IsNullOrWhiteSpace(txtPatternName.Text))
                {
                    DSP_MSG("E102", "名前", string.Empty, string.Empty, string.Empty, string.Empty);
                    txtPatternName.Focus();
                }

            }
        }

        private void ucPattern_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (!string.IsNullOrWhiteSpace(ucPattern.UC_Code))
                {
                    if (ucPattern.SelectData())
                    {
                        ucPattern.Disabled();
                        if (FormMode == "修正")
                        {
                            SetMaker_Data(ucPattern.UC_Code);
                            txtPatternName.Enabled = true;
                            //dgvdelivery.Enabled = true;
                            //dgvKamoku.Enabled = true;
                            dgvdelivery.ReadOnly = false;
                            dgvKamoku.ReadOnly = false;
                            dgvSuuryoHenKan.ReadOnly = false;
                            dgvKamoku.Columns[0].ReadOnly = true;
                            dgvdelivery.Columns[0].ReadOnly = true;
                            dgvSuuryoHenKan.Columns[0].ReadOnly = true;
                        }
                        if (FormMode == "削除" || FormMode == "照会")
                        {
                            SetMaker_Data(ucPattern.UC_Code);
                        }
                        if (FormMode == "新規")
                        {
                            DSP_MSG("E107", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                            ucPattern.Enable();
                            ucPattern.SetFocus();
                            ucPattern.UC_SearchEnable = false;
                        }

                    }
                    else
                    {
                        if (FormMode == "新規")
                        {
                            txtPatternName.Focus();
                            SendKeys.Send("+{TAB}");
                        }
                        else
                        {
                            DSP_MSG("E101", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                            ucPattern.SetFocus();
                        }
                    }
                }
                else
                {
                    DSP_MSG("E102", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                    ucPattern.SetFocus();
                }
            }
        }

        #endregion

        /// <summary>
        /// show popup by SearchMode
        /// </summary>
        private void F9()
        {
            if (SearchMode == 1)
                ucPattern.ShowSearchForm();
            else if (SearchMode == 2)
                ucCopyPattern.ShowSearchForm();
        }
    }
}
