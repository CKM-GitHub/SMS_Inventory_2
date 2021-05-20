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
using SMS.PopUp;
using System.Collections;
using SMS.CustomControls;

namespace SMS.Inventory
{
    public partial class frmPSKS0102M : SMS_Base
    {
        Login_Info loginInfo;
        DataTable dtTemp;
        M_MakerBrand_Entity mmbe;
        M_SHIIRESAKI_Entity mse;
        M_Brand_Entity mbe;
        M_MakerZaiko_H_Entity mzhe;
        PSKS0102M_BL psks0102mbl;
        L_Log_Entity lle;
        int SearchMode = 1;
        bool lastKeyIsEnter = false;

        #region constructor
        /// <summary>
        /// default
        /// </summary>
        public frmPSKS0102M()
        {
            InitializeComponent();
        }

        public frmPSKS0102M(Login_Info loginInfo)
            : base(loginInfo)
        {
            InitializeComponent();
            psks0102mbl = new PSKS0102M_BL();
            this.loginInfo = loginInfo;
            //DisableFunction();
            lblMode.Visible = false;
        }

        #endregion

        //private void DisableFunction()
        //{
        //    FunctionButtonDisabled(2);
        //    FunctionButtonDisabled(3);
        //    FunctionButtonDisabled(4);
        //    FunctionButtonDisabled(5);
        //    FunctionButtonDisabled(7);
        //    FunctionButtonDisabled(8);
        //    FunctionButtonDisabled(10);
        //    FunctionButtonDisabled(11);
        //    FunctionButtonDisabled(12);
        //}


        /// <summary>
        /// form load
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmPSKS0102M_Load(object sender, EventArgs e)
        {
            FormName = "メーカー・ブランド設定マスタ";        
            SetGridDesign();
            //CreateTempTable();
            FunctionButtonDisabled(2);
            FunctionButtonDisabled(3);
            FunctionButtonDisabled(4);
            FunctionButtonDisabled(5);
            FunctionButtonDisabled(7);
            FunctionButtonDisabled(8);
            FunctionButtonDisabled(10);
            FunctionButtonDisabled(11);
            FunctionButtonDisabled(12);
            ucMaker.SetFocus();
        }

        private void CreateTempTable()
        {
            dtTemp = new DataTable();
            dtTemp.Columns.Add("No", typeof(string));
            dtTemp.Columns.Add("MakerCD", typeof(string));
            dtTemp.Columns.Add("MakerName", typeof(string));
            dtTemp.Columns.Add("BrandCD", typeof(string));
            dtTemp.Columns.Add("BrandName", typeof(string));
            dtTemp.Columns.Add("DataSourseMakerCD", typeof(string));
            dtTemp.Columns.Add("DataMakerName", typeof(string));
            dtTemp.Columns.Add("PatternCD", typeof(string));
            dtTemp.Columns.Add("PatternName", typeof(string));

            dgvM_MakerZaiko.DataSource = dtTemp;
        }

        /// <summary>
        /// Set Grid Design in form load
        /// </summary>
        private void SetGridDesign()
        {
            string readonlyColumn = "colno,colMakerName,colBrandName,colDataName,colPatternName";
            if (ucMaker.SelectData())
            {
                dgvM_MakerZaiko.Columns["colMakerCd"].ReadOnly = true;
                dgvM_MakerZaiko.Columns["btnMaker"].ReadOnly = true;
            }
            else
            {
                dgvM_MakerZaiko.Columns["colMakerCd"].ReadOnly = false;
                dgvM_MakerZaiko.Columns["btnMaker"].ReadOnly = false;
            }

            if (ucBrand.SelectData())
            {
                dgvM_MakerZaiko.Columns["colBrandCD"].ReadOnly = true;
                dgvM_MakerZaiko.Columns["btnBrand"].ReadOnly = true;
            }
            else
            {
                dgvM_MakerZaiko.Columns["colBrandCD"].ReadOnly = false;
                dgvM_MakerZaiko.Columns["btnBrand"].ReadOnly = false;
            }

            dgvM_MakerZaiko.DisabledColumn(readonlyColumn);
            dgvM_MakerZaiko.DisabledSorting();
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
                case 2:
                case 3:
                case 4:
                case 6:
                    CreateTempTable();
                    ucMaker.Enable();
                    ucBrand.Enable();
                    uc_DatasourceMaker.Enable();
                    ucMaker.Clear();
                    ucBrand.Clear();
                    uc_DatasourceMaker.Clear();
                    FunctionButtonDisabled(12);
                    break;
                case 11:
                    break;
                case 9:
                    F9();
                    break;
                case 12:
                    F12();

                    break;
            }

        }

        #region Searching
        private void button11_Click(object sender, EventArgs e)
        {
            F11();
        }

        private void F11()
        {
            //bool makerExists = ucMaker.SelectData();
            //***BugNo22***
            bool brandExists = ucBrand.SelectData();
            if (!string.IsNullOrWhiteSpace(ucBrand.UC_Code) && !brandExists)
            {
                DSP_MSG("E101", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                ucBrand.SetFocus();
                return;
            }

            SetGridDesign();
            FunctionButtonEnabled(12);
            mmbe = new M_MakerBrand_Entity();
            mmbe.MakerCD = ucMaker.UC_Code;
            mmbe.BrandCD = ucBrand.UC_Code;
            mmbe.DatasourceMakerCD = uc_DatasourceMaker.UC_Code;

            DataTable dt = psks0102mbl.M_MakerBrand_Select(mmbe);
            if (dt.Rows.Count > 0)
            {

                dt.Rows.Add();
                dt.Rows[dt.Rows.Count - 1]["No"] = dt.Rows.Count.ToString();
                dgvM_MakerZaiko.DataSource = dtTemp = dt;
            }
            else
            {
                CreateTempTable();

            }

            ucMaker.Disabled();
            ucBrand.Disabled();
            uc_DatasourceMaker.Disabled();
        }

        #endregion

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvM_MakerZaiko_Paint(object sender, PaintEventArgs e)
        {
            string[] monthes = { "No", "メーカー(仕入先)", "ブランド", "データ元(仕入先)", "取込パターン" };
            for (int j = 0; j < 15; )
            {
                Rectangle r1 = this.dgvM_MakerZaiko.GetCellDisplayRectangle(j, -1, true);
                int w1 = this.dgvM_MakerZaiko.GetCellDisplayRectangle(j + 1, -1, true).Width;
                int w2 = this.dgvM_MakerZaiko.GetCellDisplayRectangle(j + 2, -1, true).Width;
                r1.X += 2;
                r1.Y += 1;
                r1.Width = r1.Width + w1 + w2 - 3;
                r1.Height = r1.Height - 2;

                e.Graphics.FillRectangle(new SolidBrush(this.dgvM_MakerZaiko.ColumnHeadersDefaultCellStyle.BackColor), r1);
                StringFormat format = new StringFormat();
                format.LineAlignment = StringAlignment.Center;
                e.Graphics.DrawString(monthes[j / 3],
                this.dgvM_MakerZaiko.ColumnHeadersDefaultCellStyle.Font,
                new SolidBrush(this.dgvM_MakerZaiko.ColumnHeadersDefaultCellStyle.ForeColor),
                r1,
                format);
                j += 3;
            }
        }

        /// <summary>
        /// Change grid column width
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvM_MakerZaiko_ColumnWidthChanged(object sender, DataGridViewColumnEventArgs e)
        {
            Rectangle rtHeader = this.dgvM_MakerZaiko.DisplayRectangle;
            this.dgvM_MakerZaiko.Invalidate(rtHeader);
        }

        private void dgvM_MakerZaiko_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex == -1 && e.ColumnIndex > -1)
            {
                Rectangle r2 = e.CellBounds;
                r2.Y += e.CellBounds.Height;
                r2.Height = e.CellBounds.Height;
                e.PaintBackground(r2, true);
                e.PaintContent(r2);
                e.Handled = true;
            }
        }

        /// <summary>
        /// Control grid column
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgvM_MakerZaiko_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            e.Control.Validating -= new CancelEventHandler(MakerCD_Validating);
            e.Control.Validating -= new CancelEventHandler(BrandCD_Validating);
            e.Control.Validating -= new CancelEventHandler(DataCD_Validating);
            e.Control.Validating -= new CancelEventHandler(PatternCD_Validating);            

            e.Control.PreviewKeyDown -= dgvM_MakerZaiko_PreviewKeyDown;

            if (dgvM_MakerZaiko.CurrentCell.ColumnIndex == dgvM_MakerZaiko.Columns["colMakerCD"].Index)
            {
                e.Control.Validating += new CancelEventHandler(MakerCD_Validating);
                e.Control.PreviewKeyDown += dgvM_MakerZaiko_PreviewKeyDown;
            }
            if (dgvM_MakerZaiko.CurrentCell.ColumnIndex == dgvM_MakerZaiko.Columns["colBrandCD"].Index)
            {
                e.Control.Validating += new CancelEventHandler(BrandCD_Validating);
                e.Control.PreviewKeyDown += dgvM_MakerZaiko_PreviewKeyDown;
            }
            if (dgvM_MakerZaiko.CurrentCell.ColumnIndex == dgvM_MakerZaiko.Columns["colData"].Index)
            {
                e.Control.Validating += new CancelEventHandler(DataCD_Validating);
                e.Control.PreviewKeyDown += dgvM_MakerZaiko_PreviewKeyDown;
            }
            if (dgvM_MakerZaiko.CurrentCell.ColumnIndex == dgvM_MakerZaiko.Columns["colPatternCD"].Index)
            {
                e.Control.Validating += new CancelEventHandler(PatternCD_Validating);
                e.Control.PreviewKeyDown += dgvM_MakerZaiko_PreviewKeyDown;
            }

            e.Control.KeyPress += new KeyPressEventHandler(txt_KeyPress);
        }

        /// <summary>
        /// Add new row to grid
        /// </summary>
        /// <param name="row"></param>
        private void CheckRowAdd(DataGridViewRow row)
        {
            if (row.Index == dgvM_MakerZaiko.Rows.Count - 1)
            {
                if (!string.IsNullOrWhiteSpace(row.Cells[dgvM_MakerZaiko.Columns["colMakerCD"].Index].Value.ToString())
                    && !string.IsNullOrWhiteSpace(row.Cells[dgvM_MakerZaiko.Columns["colBrandCD"].Index].Value.ToString())
                    && !string.IsNullOrWhiteSpace(row.Cells[dgvM_MakerZaiko.Columns["colData"].Index].Value.ToString())
                    && !string.IsNullOrWhiteSpace(row.Cells[dgvM_MakerZaiko.Columns["colPatternCD"].Index].Value.ToString())
                    )
                {
                    dtTemp.Rows.Add();
                    dtTemp.Rows[dtTemp.Rows.Count - 1]["No"] = dtTemp.Rows.Count.ToString();
                }
            }
        }

        private void dgvM_MakerZaiko_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            var senderGrid = (DataGridView)sender;
            if (e.RowIndex >= 0)
            {
                var row = this.dgvM_MakerZaiko.Rows[e.RowIndex];
                if (senderGrid.Columns[e.ColumnIndex] is DataGridViewButtonColumn)
                {
                    if (senderGrid.Columns[e.ColumnIndex].ReadOnly == false)
                    {
                        if (dgvM_MakerZaiko.Columns["btnMaker"].Index == e.ColumnIndex)
                        {
                            frmMaker_List ms = new frmMaker_List();
                            ms.ShowDialog();

                            if (!string.IsNullOrWhiteSpace(ms.MakerCD))
                            {
                                row.Cells[dgvM_MakerZaiko.Columns[e.ColumnIndex - 1].Index].Value = ms.MakerCD;
                                row.Cells[dgvM_MakerZaiko.Columns[e.ColumnIndex + 1].Index].Value = ms.MakerName;

                                if (!string.IsNullOrWhiteSpace(ucBrand.UC_Code))
                                {
                                    row.Cells[dgvM_MakerZaiko.Columns["colBrandCD"].Index].Value = ucBrand.UC_Code;
                                    row.Cells[dgvM_MakerZaiko.Columns["colBrandName"].Index].Value = ucBrand.UC_Name;
                                }
                            
                                CheckRowAdd(row);// to add new row
                            }

                        }

                        if (dgvM_MakerZaiko.Columns["btnBrand"].Index == e.ColumnIndex)
                        {
                            frmBrand_List frmbrand = new frmBrand_List();
                            frmbrand.ShowDialog();

                            if (!string.IsNullOrWhiteSpace(frmbrand.brandCD))
                            {
                                row.Cells[dgvM_MakerZaiko.Columns["colBrandCD"].Index].Value = frmbrand.brandCD;
                                row.Cells[dgvM_MakerZaiko.Columns["colBrandName"].Index].Value = frmbrand.brandName;

                                if (!string.IsNullOrWhiteSpace(ucMaker.UC_Code))
                                {
                                    row.Cells[dgvM_MakerZaiko.Columns["colMakerCD"].Index].Value = ucMaker.UC_Code;
                                    row.Cells[dgvM_MakerZaiko.Columns["colMakerName"].Index].Value = ucMaker.UC_Name;
                                }

                                CheckRowAdd(row);//to add new row
                            }

                        }

                        if (dgvM_MakerZaiko.Columns["btnData"].Index == e.ColumnIndex)
                        {
                            frmMaker_List ms = new frmMaker_List();
                            ms.ShowDialog();

                            if (!string.IsNullOrWhiteSpace(ms.MakerCD))
                            {
                                row.Cells[dgvM_MakerZaiko.Columns["colData"].Index].Value = ms.MakerCD;
                                row.Cells[dgvM_MakerZaiko.Columns["colDataName"].Index].Value = ms.MakerName;

                                if (!string.IsNullOrWhiteSpace(ucMaker.UC_Code))
                                {
                                    row.Cells[dgvM_MakerZaiko.Columns["colMakerCD"].Index].Value = ucMaker.UC_Code;
                                    row.Cells[dgvM_MakerZaiko.Columns["colMakerName"].Index].Value = ucMaker.UC_Name;
                                }

                                if (!string.IsNullOrWhiteSpace(ucBrand.UC_Code))
                                {
                                    row.Cells[dgvM_MakerZaiko.Columns["colBrandCD"].Index].Value = ucBrand.UC_Code;
                                    row.Cells[dgvM_MakerZaiko.Columns["colBrandName"].Index].Value = ucBrand.UC_Name;
                                }
                                    
                               CheckRowAdd(row);//to add new row
                            }
                         
                        }

                        if (dgvM_MakerZaiko.Columns["btnPattern"].Index == e.ColumnIndex)
                        {
                            frmPattern_List mzkh = new frmPattern_List();
                            mzkh.ShowDialog();

                            if (!string.IsNullOrWhiteSpace(mzkh.PatternCD))
                            {
                                row.Cells[dgvM_MakerZaiko.Columns["colPatternCD"].Index].Value = mzkh.PatternCD;
                                row.Cells[dgvM_MakerZaiko.Columns["colPatternName"].Index].Value = mzkh.PatternName;

                                if (!string.IsNullOrWhiteSpace(ucMaker.UC_Code))
                                {
                                    row.Cells[dgvM_MakerZaiko.Columns["colMakerCD"].Index].Value = ucMaker.UC_Code;
                                    row.Cells[dgvM_MakerZaiko.Columns["colMakerName"].Index].Value = ucMaker.UC_Name;
                                }

                                if (!string.IsNullOrWhiteSpace(ucBrand.UC_Code))
                                {
                                    row.Cells[dgvM_MakerZaiko.Columns["colBrandCD"].Index].Value = ucBrand.UC_Code;
                                    row.Cells[dgvM_MakerZaiko.Columns["colBrandName"].Index].Value = ucBrand.UC_Name;
                                }

                                CheckRowAdd(row);//to add new row
                            }
                           
                        }
                    }
                }
            }
        }

        #region Save adding or editing records
        private void F12()
        {
            string confirmMessageID = "Q101";
            string msg = DSP_MSG(confirmMessageID, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);

            if (msg.Equals("1"))
            {
                if (ErrorCheck())
                {
                    Save();
                    functionClick(6);
                    ucMaker.SetFocus();
                }
            }
        }

        private bool Save()
        {
            string datetime = DateTime.Now.ToString();

            mmbe = new M_MakerBrand_Entity();
            mmbe.KAICD = loginInfo.CompanyCode;
            mmbe.MakerCD = ucMaker.UC_Code;
            mmbe.BrandCD = ucBrand.UC_Code;
            mmbe.DatasourceMakerCD = uc_DatasourceMaker.UC_Code;
            mmbe.InsertOperator = loginInfo.OperatorCode;
            mmbe.InsertDateTime = datetime;

            lle = new L_Log_Entity();
            lle.OperateDate = DateTime.Now.ToString();
            lle.Program = "PSKS0102M";
            lle.PC = loginInfo.PcName;
            lle.InsertOperator = loginInfo.OperatorCode;
            lle.OperateMode = "新規";
            lle.KeyItem = ucMaker.UC_Code + " " + ucBrand.UC_Name;
            mmbe.dt1 = dtTemp;
            return psks0102mbl.M_MakerBrand_Insert_Xml(mmbe, lle);
        }

        #endregion

        /// <summary>
        /// Error Check function
        /// </summary>
        /// <returns></returns>
        private bool ErrorCheck()
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("colMakerCD,MakerCD", "MakerCD");
            dic.Add("colBrandCD,BrandCD", "BrandCD");
            dic.Add("colData,DataSourseMakerCD", "DataSourseMakerCD");
            dic.Add("colPatternCD,PatternCD", "PatternCD");

            if (!NullCheckForGrid(dgvM_MakerZaiko, dtTemp, dic))
            {
                return false;
            }

            dic = new Dictionary<string, string>();
            dic.Add("MakerCD","VC_SHIIRESAKI,M_SHIIRESAKI");
            dic.Add("BrandCD","nc_brand,m_brand");
            dic.Add("DataSourseMakerCD", "VC_SHIIRESAKI,M_SHIIRESAKI");
            dic.Add("PatternCD","PatternCD,M_MakerZaiko_H" );
            if (!DataExistsCheckForGrid(dgvM_MakerZaiko, dtTemp, dic))
                return false;

            DataTable distinct = new DataTable();
            DataTable dt = new DataTable();
            DataTable dtDuplicate = new DataTable();
            DataRow[] dr = dtTemp.Select("MakerCD <> '' OR MakerCD IS NOT  NULL AND BrandCD IS NOT  NULL ");
            if (dr.Count() > 0)
            {
                dt = dr.CopyToDataTable();

            }

            foreach (DataRow drr in dt.Rows)
            {
                dr = dt.Select("BrandCD = '" + drr["BrandCD"] + "' AND MakerCD = '" + drr["MakerCD"] + "'");
                if (dr.Count() > 1)
                {
                    DSP_MSG("E105", "メーカー値", string.Empty, string.Empty, string.Empty, string.Empty);
                    dgvM_MakerZaiko.Select();
                    dgvM_MakerZaiko.CurrentCell = dgvM_MakerZaiko[dgvM_MakerZaiko.Columns["colBrandCD"].Index, Convert.ToInt16(dr[0]["No"].ToString()) - 1];//move cursor to duplicate cell
                    return false;
                }
            }
            
            return true;
        }

        //private DataTable SelectDuplicateRow(DataTable dtTemp)
        //{
        //    ArrayList UniqueRecords = new ArrayList();
        //    ArrayList DuplicateRecords = new ArrayList();
        //    foreach (DataRow dRow in dtTemp.Rows)
        //    {
        //        if (UniqueRecords.Contains(dRow["BrandCD"]))
        //            DuplicateRecords.Add(dRow);
        //    }


        //    return dtTemp;
        //}

        #region Key Down
        private void ucMaker_KeyDown(object sender, KeyEventArgs e)
        {
            if ((e.KeyCode == Keys.Enter) && (!string.IsNullOrWhiteSpace(ucMaker.UC_Code)))
            {
                if (!ucMaker.SelectData())
                {
                    DSP_MSG("E101", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                }
            }
        }

        private void uc_DatasourceMaker_KeyDown(object sender, KeyEventArgs e)
        {
            if ((e.KeyCode == Keys.Enter) && (!string.IsNullOrWhiteSpace(uc_DatasourceMaker.UC_Code)))
            {
                if (!uc_DatasourceMaker.SelectData())
                {
                    DSP_MSG("E101", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                }
            }
        }

        private void ucBrand_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if ((string.IsNullOrWhiteSpace(ucBrand.UC_Code)) || ucBrand.SelectData())
                {
                    F11();
                }
                else
                {
                    DSP_MSG("E101", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                }
            }
        }

        private void dgvM_MakerZaiko_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                lastKeyIsEnter = true;
            else
                lastKeyIsEnter = false;
        }

        private void frmPSKS0102M_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F11)
            {
                F11();
            }
        }

        #endregion

        #region Validating
        private void MakerCD_Validating(object sender, CancelEventArgs e)
        {
            TextBox txt = sender as TextBox;
            int row = dgvM_MakerZaiko.CurrentCell.RowIndex;

            if (string.IsNullOrWhiteSpace(txt.Text))
            {
                dgvM_MakerZaiko.Rows[row].Cells["colMakerName"].Value = string.Empty;
                dgvM_MakerZaiko.Rows[row].Cells["colBrandCD"].Value = string.Empty;
                dgvM_MakerZaiko.Rows[row].Cells["colBrandName"].Value = string.Empty;
                dgvM_MakerZaiko.Rows[row].Cells["colData"].Value = string.Empty;
                dgvM_MakerZaiko.Rows[row].Cells["colDataName"].Value = string.Empty;
                dgvM_MakerZaiko.Rows[row].Cells["colPatternCD"].Value = string.Empty;
                dgvM_MakerZaiko.Rows[row].Cells["colPatternName"].Value = string.Empty;
                return;
            } 

            mse = new M_SHIIRESAKI_Entity();
            mse.MakerCD = txt.Text;
            DataTable dt = psks0102mbl.M_Shiiresaki_Select(mse);
            if (dt.Rows.Count > 0)
            {
                dgvM_MakerZaiko.Rows[row].Cells["colMakerName"].Value = dt.Rows[0]["VM_SHIIRESAKI"].ToString();

                if (!string.IsNullOrWhiteSpace(ucBrand.UC_Code))
                {
                    dgvM_MakerZaiko.Rows[row].Cells[dgvM_MakerZaiko.Columns["colBrandCD"].Index].Value = ucBrand.UC_Code;
                    dgvM_MakerZaiko.Rows[row].Cells[dgvM_MakerZaiko.Columns["colBrandName"].Index].Value = ucBrand.UC_Name;
                }

                CheckRowAdd(dgvM_MakerZaiko.Rows[row]);
            }
            else
            {
                DSP_MSG("E101", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                dgvM_MakerZaiko.CurrentCell = dgvM_MakerZaiko[dgvM_MakerZaiko.Columns["colMakerCD"].Index, row];
                if (lastKeyIsEnter)
                    SendKeys.Send("+{TAB}");
                else
                    e.Cancel = true;
            }
        }

        private void BrandCD_Validating(object sender, CancelEventArgs e)
        {
            TextBox txt = sender as TextBox;
            int row = dgvM_MakerZaiko.CurrentCell.RowIndex;
            if (string.IsNullOrWhiteSpace(txt.Text))
            {
                dgvM_MakerZaiko.Rows[row].Cells["colBrandName"].Value = string.Empty;
                dgvM_MakerZaiko.Rows[row].Cells["colMakerCd"].Value = string.Empty;
                dgvM_MakerZaiko.Rows[row].Cells["colMakerName"].Value = string.Empty;
                dgvM_MakerZaiko.Rows[row].Cells["colData"].Value = string.Empty;
                dgvM_MakerZaiko.Rows[row].Cells["colDataName"].Value = string.Empty;
                dgvM_MakerZaiko.Rows[row].Cells["colPatternCD"].Value = string.Empty;
                dgvM_MakerZaiko.Rows[row].Cells["colPatternName"].Value = string.Empty;
                return;
            }

            mbe = new M_Brand_Entity();
            mbe.BrandCD = txt.Text;

            DataTable dt = psks0102mbl.M_Brand_Select(mbe);
            if (dt.Rows.Count > 0)
            {
                dgvM_MakerZaiko.Rows[row].Cells["colBrandName"].Value = dt.Rows[0]["vm_brand"].ToString();

                if (!string.IsNullOrWhiteSpace(ucMaker.UC_Code))
                {
                    dgvM_MakerZaiko.Rows[row].Cells[dgvM_MakerZaiko.Columns["colMakerCD"].Index].Value = ucMaker.UC_Code;
                    dgvM_MakerZaiko.Rows[row].Cells[dgvM_MakerZaiko.Columns["colMakerName"].Index].Value = ucMaker.UC_Name;
                }

                CheckRowAdd(dgvM_MakerZaiko.Rows[row]);
            }
            else
            {
                DSP_MSG("E101", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                dgvM_MakerZaiko.Rows[row].Cells["colBrandName"].Value = string.Empty;
                dgvM_MakerZaiko.CurrentCell = dgvM_MakerZaiko[dgvM_MakerZaiko.Columns["colBrandCD"].Index, row];
                if (lastKeyIsEnter)
                    SendKeys.Send("+{TAB}");
                else
                    e.Cancel = true;
            }
        }

        private void DataCD_Validating(object sender, CancelEventArgs e)
        {
            TextBox txt = sender as TextBox;
            int row = dgvM_MakerZaiko.CurrentCell.RowIndex;
            if (string.IsNullOrWhiteSpace(txt.Text))
            {
                
                dgvM_MakerZaiko.Rows[row].Cells["colDataName"].Value = string.Empty;
                return;
            }

            mse = new M_SHIIRESAKI_Entity();
            mse.MakerCD = txt.Text;
            DataTable dt = psks0102mbl.M_Shiiresaki_Select(mse);

            if (dt.Rows.Count > 0)
            {
                dgvM_MakerZaiko.Rows[row].Cells["colDataName"].Value = dt.Rows[0]["DataMakerName"].ToString();

                if (!string.IsNullOrWhiteSpace(ucMaker.UC_Code))
                {
                    dgvM_MakerZaiko.Rows[row].Cells[dgvM_MakerZaiko.Columns["colMakerCD"].Index].Value = ucMaker.UC_Code;
                    dgvM_MakerZaiko.Rows[row].Cells[dgvM_MakerZaiko.Columns["colMakerName"].Index].Value = ucMaker.UC_Name;
                }

                if (!string.IsNullOrWhiteSpace(ucBrand.UC_Code))
                {
                    dgvM_MakerZaiko.Rows[row].Cells[dgvM_MakerZaiko.Columns["colBrandCD"].Index].Value = ucBrand.UC_Code;
                    dgvM_MakerZaiko.Rows[row].Cells[dgvM_MakerZaiko.Columns["colBrandName"].Index].Value = ucBrand.UC_Name;
                }

                CheckRowAdd(dgvM_MakerZaiko.Rows[row]);
            }
            else
            {
                DSP_MSG("E101", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                dgvM_MakerZaiko.Rows[row].Cells["colDataName"].Value = string.Empty;
                dgvM_MakerZaiko.CurrentCell = dgvM_MakerZaiko[dgvM_MakerZaiko.Columns["colData"].Index, row];
                if (lastKeyIsEnter)
                    SendKeys.Send("+{TAB}");
                else
                    e.Cancel = true;
            }
        }

        private void PatternCD_Validating(object sender, CancelEventArgs e)
        {
            TextBox txt = sender as TextBox;
            int row = dgvM_MakerZaiko.CurrentCell.RowIndex;

            if (string.IsNullOrWhiteSpace(txt.Text))
            {
                dgvM_MakerZaiko.Rows[row].Cells["colPatternName"].Value = string.Empty;

                return;
            }

            mzhe = new M_MakerZaiko_H_Entity();
            mzhe.PatternCD = txt.Text;
            DataTable dt = psks0102mbl.M_MakerZaiko_H_Select(mzhe);

            if (dt.Rows.Count > 0)
            {
                dgvM_MakerZaiko.Rows[row].Cells["colPatternName"].Value = dt.Rows[0]["PatternName"].ToString();

                if (!string.IsNullOrWhiteSpace(ucMaker.UC_Code))
                {
                    dgvM_MakerZaiko.Rows[row].Cells[dgvM_MakerZaiko.Columns["colMakerCD"].Index].Value = ucMaker.UC_Code;
                    dgvM_MakerZaiko.Rows[row].Cells[dgvM_MakerZaiko.Columns["colMakerName"].Index].Value = ucMaker.UC_Name;
                }

                if (!string.IsNullOrWhiteSpace(ucBrand.UC_Code))
                {
                    dgvM_MakerZaiko.Rows[row].Cells[dgvM_MakerZaiko.Columns["colBrandCD"].Index].Value = ucBrand.UC_Code;
                    dgvM_MakerZaiko.Rows[row].Cells[dgvM_MakerZaiko.Columns["colBrandName"].Index].Value = ucBrand.UC_Name;
                }

                CheckRowAdd(dgvM_MakerZaiko.Rows[row]);
            }
            else
            {
                DSP_MSG("E101", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                dgvM_MakerZaiko.Rows[row].Cells["colPatternName"].Value = string.Empty;
                dgvM_MakerZaiko.CurrentCell = dgvM_MakerZaiko[dgvM_MakerZaiko.Columns["colPatternCD"].Index, row];
                if (lastKeyIsEnter)
                    SendKeys.Send("+{TAB}");
                else
                    e.Cancel = true;
            }
        }

        #endregion

        /// <summary>
        /// Key press
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void txt_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (((e.KeyChar < 48 || e.KeyChar > 57) && e.KeyChar != 8 && e.KeyChar != 45))
            {
                e.Handled = true;
            }
        }


        /// <summary>
        /// show popup by SearchMode
        /// </summary>
        private void F9()
        {
            if (SearchMode == 1)
                ucMaker.ShowSearchForm();
            else if (SearchMode == 2)
                ucBrand.ShowSearchForm();
            else if (SearchMode == 3)
                uc_DatasourceMaker.ShowSearchForm();
        }

        #region Handle for F9 Function

        /// <summary>
        ///  handle all usercontrol to show search popup
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

        private void uc_DatasourceMaker_Enter(object sender, EventArgs e)
        {
            SearchMode = 3;
            FunctionButtonEnabled(9);
        }


        #endregion

    }
}
