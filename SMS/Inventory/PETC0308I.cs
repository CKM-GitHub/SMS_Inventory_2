using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//using System.Windows.Forms;
using SMS_BL;
using SMS_Entity;
//using System.Data;
using System.IO;
using System.Diagnostics;
using ElencySolutions.CsvHelper;
using Microsoft.VisualBasic.FileIO;
using System.Globalization;
using DataTableToCSV;
using System.Windows.Forms;

namespace SMS.Inventory
{
    public partial class frmPETC0308I : SMS_Base
    {
        Login_Info loginInfo;
        PETC0308I_BL PETC0308IBL = new PETC0308I_BL();
        Base_BL baseBL = new Base_BL();
        L_Log_Entity LLogEntity;
        T_Tesuuryou_Entity Tessuryou_data;
        T_NyuukinMeisai_Entity NyuukinMeisai_data;
        string firstColHeader = string.Empty;
        string filename = string.Empty;
        string Date = string.Empty;
        string xml = string.Empty;
        string MotoCD = string.Empty;
        string KouzaDate = string.Empty;
        bool TabSkip = false;
        bool set_focus = false;
        DataGridViewCell dgcell = null;
        DataTable dtCSV, dtNyuukin = new DataTable();
        DataTable file = new DataTable();
        public frmPETC0308I()
        {
            InitializeComponent();
        }

        public frmPETC0308I(Login_Info loginInfo)
            : base(loginInfo)
        {
            InitializeComponent();
            this.loginInfo = loginInfo;
        }
        private void frmPETC0308I_Load(object sender, EventArgs e)
        {
            
            FunctionButtonDisabled(2);
            FunctionButtonDisabled(3);
            FunctionButtonDisabled(4);
            FunctionButtonDisabled(5);
            FunctionButtonDisabled(6);
            FunctionButtonDisabled(7);
            FunctionButtonDisabled(8);
            FunctionButtonDisabled(9);
            FunctionButtonDisabled(10);

            F12_Text = "出力(F12)";
            F11_Text = "取込(F11)";
            FormName = "入金手数料(Web)仕訳作成MF";
            dgvDetail.EnabledColumn("colchk");
            dgvPETC0308I.Columns["colNo"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvDetail.Columns["colNo1"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        //    dgvPETC0308I.DisabledColumn("colNo,colNyuukinMotoName1,colKouzaNyuukinDate,col_nyuukingaku");
            lblMode.Visible = false;
            dgvDetail.Gridcolor(Color.Black);
            dgvPETC0308I.Gridcolor(Color.Black);
            txtdocDate.Text = DateTime.Now.ToString("yyyy/MM/dd").ToString().Replace("-","/");
            dgvPETC0308I.EditingControlShowing += datagridview_EditingControlShowing;
            PETC0308IBL = new PETC0308I_BL();
            //var dt = PETC0308IBL.Mshiwake_title();
            var dt = PETC0308IBL.Mshiwake_title();
            for (int j = 0; j < 15; j++)
            {
                var col = (j + 1).ToString().PadLeft(2, '0');
                dgvPETC0308I.Columns["colKeihi" + col + "Input"].ValueType = typeof(string);//KeihiTitle01
                dgvPETC0308I.Columns["colKeihi" + col + "Input"].HeaderText = dt.Rows[0]["KeihiTitle" + col].ToString();
                Control[] controls = this.Controls.Find("lbl_total" + (j + 1), true);
                if (controls.Length == 1) // 0 means not found, more - there are several controls with the same name
                {
                    CKM_Controls.CKM_Label control = controls[0] as CKM_Controls.CKM_Label;
                    control.value = "0";
                }
            }

            DisableHeaderFlexible(dgvPETC0308I);
            DisableHeaderFlexible(dgvDetail);
            //    dgvDetail.VirtualMode = true;


            checkState = new Dictionary<int, bool>();
        }

        // private System.Collections.Generic.Dictionary<int, bool> checkState;

        public void DisableHeaderFlexible(DataGridView dg)
        {
            dg.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            ///dg.ColumnSortModeChanged   +=dg_ColumnSortModeChanged;
            for (int i = 0; i < dg.Columns.Count; i++)
            {
                dg.Columns[i].Resizable = DataGridViewTriState.False;
                dg.Columns[i].Resizable = DataGridViewTriState.False;
                dg.Columns[i].SortMode = DataGridViewColumnSortMode.Automatic;
            }
            foreach (DataGridViewColumn column in dg.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.Automatic;
            }
        }
        Control cntObject;  /// Make global declare

        private void datagridview_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)// Select DataGridView EditingControlShowing Event
        {
            e.Control.TextChanged += new EventHandler(textBox_TextChanged);
            cntObject = e.Control;
            cntObject.TextChanged += textBox_TextChanged;
        }

        private void textBox_TextChanged(object sender, EventArgs e)// TextBox TextChanged Event must be put When be override when not working in validation (Same with dgvcontrol columns text chaning event but not still used on this)
        {
            if (cntObject.Text != string.Empty)
            {

                cntObject.Select();

            }
        }

        private void btnAllCheckon_Click(object sender, EventArgs e)  // All Check on
        {
            if (!String.IsNullOrEmpty(txtImport.Text))
            {
                BindGrid(txtdocDate.Text);
            }
            CheckALL(dgvPETC0308I, "coltesu_chk", null, true, null);
            CheckALL(dgvDetail, "colchk", null, true, null);
        }
        public void CheckALL(DataGridView gv, string columnname, string[] CD_kouza = null, bool flg = false, string Date = null)
        {
            if (CD_kouza == null && Date == null)
            {
                foreach (DataGridViewRow row in gv.Rows)
                {
                    if (flg)
                    {
                        row.Cells[columnname].Value = true;
                    }
                    else
                    {
                        row.Cells[columnname].Value = false;
                    }
                }
            }
            else if (CD_kouza != null && Date == null)
            {
                foreach (DataGridViewRow gr in gv.Rows)
                {
                    if (gr.Cells["colKouzaNyuukinDate"].FormattedValue.ToString() == CD_kouza[1].ToString() && gr.Cells["colNyuukinMotoCD"].FormattedValue.ToString() == CD_kouza[0].ToString())
                    {
                        if (flg)
                        {
                            gr.Cells["coltesu_chk"].Value = true;
                        }
                        else
                        {
                            gr.Cells["coltesu_chk"].Value = false;
                        }
                    }
                }
            }
            else if (CD_kouza != null && Date != null)
            {
                foreach (DataGridViewRow gr in gv.Rows)
                {
                    if (gr.Cells["colKouzaNyuukinDate"].FormattedValue.ToString() == CD_kouza[1].ToString() && gr.Cells["colNyuukinMotoCD"].FormattedValue.ToString() == CD_kouza[0].ToString())
                    {
                        gr.Cells["coldocDate"].Value = Date;

                    }
                }
            }
        }
        private void btnAllCheckOff_Click(object sender, EventArgs e) //All Check Off
        {
            if (!String.IsNullOrEmpty(txtImport.Text))
            {
                BindGrid(txtdocDate.Text);
            }

            CheckALL(dgvPETC0308I, "coltesu_chk", null, false, null);
            CheckALL(dgvDetail, "colchk", null, false, null);
            foreach (DataGridViewRow gr in dgvPETC0308I.Rows)
            {
                gr.Cells["col_nyuukingaku"].Value = "0";
            }
        }

        private void dgvDetail_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            MaintainCellClick_detail(sender, e);
        }

        private void dgvDetail_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            MaintainCellClick_detail(sender, e);

        }

        private void MaintainCellClick_detail(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
                if ((sender as DataGridView).CurrentCell is DataGridViewCheckBoxCell)
                {

                    if ((Convert.ToBoolean(dgvDetail.Rows[e.RowIndex].Cells[e.ColumnIndex].EditedFormattedValue) == true))
                    {
                        MotoCD = dgvDetail.Rows[e.RowIndex].Cells["colNyuukinMotoCD1"].Value.ToString();
                        KouzaDate = dgvDetail.Rows[e.RowIndex].Cells["colKouzaNyuukinDate1"].Value.ToString();
                        string[] moto_kouza = { MotoCD, KouzaDate };
                        CheckALL(dgvPETC0308I, "blur", moto_kouza, true);

                    }
                    else
                    {
                        // this.dgvPETC0302I.Refresh();
                        bool flg = false;
                        foreach (DataGridViewRow row1 in dgvDetail.Rows)
                        {
                            if (row1.Cells["colNyuukinMotoCD1"].Value.ToString() == MotoCD && row1.Cells["colKouzaNyuukinDate1"].Value.ToString() == KouzaDate)
                            {

                                DataGridViewCheckBoxCell cell1 = row1.Cells["colchk"] as DataGridViewCheckBoxCell;
                                bool bData1 = (bool)cell1.EditedFormattedValue;

                                if (bData1 == true)
                                {
                                    flg = true;
                                    break;
                                }
                            }
                        }
                        if (!flg)
                        {
                            MotoCD = dgvDetail.Rows[e.RowIndex].Cells["colNyuukinMotoCD1"].Value.ToString();
                            KouzaDate = dgvDetail.Rows[e.RowIndex].Cells["colKouzaNyuukinDate1"].Value.ToString();
                            string[] moto_kouza = { MotoCD, KouzaDate };
                            CheckALL(dgvPETC0308I, "blur", moto_kouza, false);
                        }

                    }


                    MotoCD = dgvDetail.Rows[e.RowIndex].Cells["colNyuukinMotoCD1"].Value.ToString();
                    KouzaDate = dgvDetail.Rows[e.RowIndex].Cells["colKouzaNyuukinDate1"].Value.ToString();
                    int res = 0;
                    foreach (DataGridViewRow row1 in dgvDetail.Rows)
                    {
                        if (row1.Cells["colNyuukinMotoCD1"].Value.ToString() == MotoCD && row1.Cells["colKouzaNyuukinDate1"].Value.ToString() == KouzaDate)
                        {
                            if ((Convert.ToBoolean(row1.Cells["colchk"].EditedFormattedValue) == true))
                            {
                                res += Convert.ToInt32(row1.Cells["colNyuukinGaku"].EditedFormattedValue.ToString().Replace(",", ""));
                            }
                        }
                    }
                    foreach (DataGridViewRow row1 in dgvPETC0308I.Rows)
                    {
                        if (row1.Cells["colNyuukinMotoCD"].Value.ToString() == MotoCD && row1.Cells["colKouzaNyuukinDate"].Value.ToString() == KouzaDate)
                        {
                            row1.Cells["col_nyuukingaku"].Value = res.ToString(); 
                          //  row1.Cells["col_nyuukingaku"].Value = CommaSeparate(res.ToString()); // repaired by Ptk 2020/01/31  Bcox of already control on Format GridView

                        }
                    }
                }
        }

        public static object CommaSeparate(string value)  // return comma separated values
        {
            object j = Round_Off(value).ToString("#,##0");
            return j;
        }
        private void dgvDetail_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dgv_CellClick(object sender, DataGridViewCellEventArgs e)
        {



            //if (dgvDetail.CurrentRow.Cells["colchk"].Value != null && (bool)dgvDetail.CurrentRow.Cells["colchk"].Value)
            //{
            //    dgvDetail.CurrentRow.Cells["colchk"].Value = false;
            //    dgvDetail.CurrentRow.Cells["colchk"].Value = null;
            //}
            //else if (dgvDetail.CurrentRow.Cells["colchk"].Value == null)
            //{
            //    dgvDetail.CurrentRow.Cells["colchk"].Value = true;
            //}
        }

        private void dgvDetail_CellValidated(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dgvDetail_Click(object sender, EventArgs e)
        {

        }

        private void dgvDetail_Validated(object sender, EventArgs e)
        {

        }

        private void dgvDetail_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex >= 0)
                if ((sender as DataGridView).CurrentCell is DataGridViewCheckBoxCell)
                {

                    if ((Convert.ToBoolean(dgvDetail.Rows[e.RowIndex].Cells[e.ColumnIndex].EditedFormattedValue) == true))
                    {
                        MotoCD = dgvDetail.Rows[e.RowIndex].Cells["colNyuukinMotoCD1"].Value.ToString();
                        KouzaDate = dgvDetail.Rows[e.RowIndex].Cells["colKouzaNyuukinDate1"].Value.ToString();
                        string[] moto_kouza = { MotoCD, KouzaDate };
                        CheckALL(dgvPETC0308I, "blur", moto_kouza, true);

                    }
                    else
                    {
                        // this.dgvPETC0302I.Refresh();
                        bool flg = false;
                        foreach (DataGridViewRow row1 in dgvDetail.Rows)
                        {
                            if (row1.Cells["colNyuukinMotoCD1"].Value.ToString() == MotoCD && row1.Cells["colKouzaNyuukinDate1"].Value.ToString() == KouzaDate)
                            {




                                DataGridViewCheckBoxCell cell1 = row1.Cells["colchk"] as DataGridViewCheckBoxCell;
                                bool bData1 = (bool)cell1.EditedFormattedValue;

                                if (bData1 == true)
                                {
                                    flg = true;
                                    break;
                                }
                            }
                        }
                        if (!flg)
                        {
                            MotoCD = dgvDetail.Rows[e.RowIndex].Cells["colNyuukinMotoCD1"].Value.ToString();
                            KouzaDate = dgvDetail.Rows[e.RowIndex].Cells["colKouzaNyuukinDate1"].Value.ToString();
                            string[] moto_kouza = { MotoCD, KouzaDate };
                            CheckALL(dgvPETC0308I, "blur", moto_kouza, false);
                        }

                    }
                }
        }

        private void dgvDetail_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dgvPETC0308I_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            //if (this.dgvPETC0302I.IsCurrentCellDirty)
            //{
            //    dgvPETC0302I.CommitEdit(DataGridViewDataErrorContexts.Commit);
            //}
        }

        private void dgvPETC0308I_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {






        }
        private void btnFolder_Click(object sender, EventArgs e)
        {
            DataTable dtPath = new DataTable();
            dtPath = PETC0308IBL.GetFolder();
            string folderPath = string.Empty;
            folderPath = dtPath.Rows[0]["Char2"].ToString();
            if (!string.IsNullOrWhiteSpace(folderPath))
            {
                if (!Directory.Exists(folderPath))
                {
                    Directory.CreateDirectory(folderPath);
                }
                OpenFileDialog op = new OpenFileDialog();
                op.InitialDirectory = folderPath;
                op.RestoreDirectory = true;
                op.Filter = "CSV|*.csv";
                if (op.ShowDialog() == DialogResult.OK)
                {
                    txtImport.Text = op.FileName;
                }
            }
        }
        public override void functionClick(int btnIndex)
        {
            switch (btnIndex)
            {
                case 1:
                    break;

                case 11:
                    F11();
                    break;

                case 12:
                    F12();
                    break;
            }
        }
        private void F12()
        {
            if (ErrorCheckForGrid())
            {
                if (DSP_MSG("Q203", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty).Equals("1"))
                {
                    //テーブル転送仕様Ｃ, テーブル転送仕様Ｄ
                    if (PETC0308IBL.Update_Tesuyou_Nyuukin(Tesuyou_dt(), Nyuukin_dt()))
                    {
                        //出力転送表01, 出力転送表02
                        PETC0308IBL = new PETC0308I_BL();
                        DataTable dtexp = new DataTable();
                        NyuukinMeisai_data = new T_NyuukinMeisai_Entity();
                        NyuukinMeisai_data.OperatorCD = loginInfo.OperatorCode;
                        Export_CSV(PETC0308IBL.NyuukinDate_NyuukinMoto_Select(NyuukinMeisai_data.OperatorCD, "1"), PETC0308IBL.PETC0302I_CSVExport1(NyuukinMeisai_data, "PETC0308I_ExportCSV"));
                    }
                }
            }
        }
        public void Export_CSV(DataTable dtdgv, DataTable dtselect)
        {
            for (int h = 0; h < dtdgv.Rows.Count; h++)
            {
                if (!string.IsNullOrWhiteSpace(dtdgv.Rows[h]["KouzaNyuukinDate"].ToString()) && !string.IsNullOrWhiteSpace(dtdgv.Rows[h]["NyuukinMotoCD"].ToString()))
                {
                    if (h == 14)
                    {
                    }
                    if (dtselect.Select("NyuukinMotoCD = '" + dtdgv.Rows[h]["NyuukinMotoCD"].ToString() + "' and KouZadate = '" + dtdgv.Rows[h]["KouzaNyuukinDate"].ToString() + "'").CopyToDataTable().Rows.Count > 0)
                    {
                        var f = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                        var g = f.Split(' ').First().ToString().Replace("/", "") + "_" + f.Split(' ').Last().Replace(":", "");
                        var dtout = dtselect.Select("NyuukinMotoCD = '" + dtdgv.Rows[h]["NyuukinMotoCD"].ToString() + "' and KouZadate = '" + dtdgv.Rows[h]["KouzaNyuukinDate"].ToString() + "'").CopyToDataTable();
                        var filename = g + "_" + loginInfo.OperatorCode + "_" + dtdgv.Rows[h]["KouzaNyuukinDate"].ToString().Replace("/", "") + "_" + dtdgv.Rows[h]["NyuukinMotoName"].ToString() + ".CSV";
                        output_csv(dtout, filename);
                    }
                }
            }
        }
        public DataTable Nyuukin_dt()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("OperatorCD");
            dt.Columns.Add("colchk");
            dt.Columns.Add("DetaKBN");
            dt.Columns.Add("colNyuukinMotoCD1");
            dt.Columns.Add("colKouzaNyuukinDate1");
            dt.Columns.Add("colDetailNo");

            foreach (DataGridViewRow gr in dgvDetail.Rows)
            {
                dt.Rows.Add();
                dt.Rows[gr.Index]["OperatorCD"] = loginInfo.OperatorCode;
                dt.Rows[gr.Index]["colchk"] = Convert.ToBoolean(gr.Cells["colchk"].EditedFormattedValue) ? "0" : "1";
                dt.Rows[gr.Index]["DetaKBN"] = "1";
                dt.Rows[gr.Index]["colNyuukinMotoCD1"] = gr.Cells["colNyuukinMotoCD1"].Value;
                dt.Rows[gr.Index]["colKouzaNyuukinDate1"] = gr.Cells["colKouzaNyuukinDate1"].Value;
                dt.Rows[gr.Index]["colDetailNo"] = gr.Cells["colDetailNo"].Value;
            }
            return dt;
        }

        public DataTable Tesuyou_dt()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("OperatorCD");
            dt.Columns.Add("DetaKBN");
            dt.Columns.Add("KouzaNyuukinDate");
            dt.Columns.Add("KeijouDate");
            dt.Columns.Add("NyuukinMotoCD");
            dt.Columns.Add("SEQ");
            for (int c = 1; c <= 15; c++)
            {
                var val = (c).ToString().PadLeft(2, '0');
                dt.Columns.Add("Keihi" + val + "input");
                dt.Columns.Add("Keihi" + val + "Tax");
            }
            dt.Columns.Add("coltesu_chk");
            for (int j = 0; j < dgvPETC0308I.Rows.Count; j++)
            {
                dt.Rows.Add();
                dt.Rows[j]["OperatorCD"] = loginInfo.OperatorCode;
                dt.Rows[j]["DetaKBN"] = "1";
                dt.Rows[j]["coltesu_chk"] = Convert.ToBoolean(dgvPETC0308I.Rows[j].Cells["coltesu_chk"].EditedFormattedValue) ? "0" : "1";
                dt.Rows[j]["KeijouDate"] = dgvPETC0308I.Rows[j].Cells["coldocDate"].Value;
            }

            for (int c = 1; c <= 15; c++)
            {
                foreach (DataGridViewRow gr in dgvPETC0308I.Rows)
                {
                    var val = (c).ToString().PadLeft(2, '0');
                    dt.Rows[gr.Index]["Keihi" + val + "Tax"] = gr.Cells["colKeihi" + val + "Tax"].Value;
                    dt.Rows[gr.Index]["Keihi" + val + "input"] = gr.Cells["colKeihi" + val + "input"].Value;
                    dt.Rows[gr.Index]["NyuukinMotoCD"] = gr.Cells["colNyuukinMotoCD"].Value;
                    dt.Rows[gr.Index]["KouzaNyuukinDate"] = gr.Cells["colKouzaNyuukinDate"].Value;
                    dt.Rows[gr.Index]["SEQ"] = gr.Cells["colseq"].Value;

                }
            }
            return dt;
        }
        protected int NumberOfSheet(int totalCount, int NumberofItems)
        {
            if (totalCount == 0 || NumberofItems == 0)
            {
                return 0;
            }
            else
            {
                int res = 0 ;
                var ffff = totalCount / NumberofItems;
                var gggg = totalCount % NumberofItems;
                if (gggg > 0)
                {
                     res = ffff + 1;
                }
                else
                {
                     res = ffff;
                }
                return res;
            }
        
        }

        public DataTable Cover_KeijouDate(DataTable dt)   // Put this bcos if there is no Cells values were put, it would be covered. 
        {
            if (Tesuyou_dt().Select("coltesu_chk <> 1").Count() > 0)
            {
                foreach (DataRow dr in Tesuyou_dt().Select("coltesu_chk <> 1").CopyToDataTable().Rows)
                {
                    foreach (DataRow drow in dt.Rows)
                    {
                        if (drow["KouZadate"].ToString() == dr["KouzaNyuukinDate"].ToString() && drow["NyuukinMotoCD"].ToString() == dr["NyuukinMotoCD"].ToString())
                        {
                            drow["取引日"] = string.IsNullOrEmpty(dr["KeijouDate"].ToString()) ? "" : dr["KeijouDate"].ToString();
                        }
                    }
                }
            }
            return dt;
        }
        public void output_csv(DataTable dtout, string filename)
        {
            DataTable dtresult = CommmaSeparatedinOutput(Cover_KeijouDate(dtout), new string[] { "借方金額(円)", "貸方金額(円)" });
            dtresult.Columns.Remove("NyuukinMotoCD");
        //    dtresult.Columns.Remove("Rowno");
            dtresult.Columns.Remove("KouZadate");
            int res = 0;
            for (int i = 0; i < dtresult.Rows.Count; i++)
            {
                if (Convert.ToInt32(dtresult.Rows[i]["Rowno"].ToString()) > 0)
                {
                    break;
                }
                else
                {
                    if (i != 0)
                    {
                 
                       // dtresult.Rows[i]["借方金額(円)"] ="0";
                        dtresult.Rows[i]["借方税額"] = "0";
                       // res += Convert.ToInt32(dtresult.Rows[i]["貸方金額(円)"].ToString().Replace(",", ""));
                    } 
                    dtresult.Rows[i]["貸方税額"] = "0";
                }
            }
            dtresult.Rows[0]["借方税額"] = "0";
           // dtresult.Rows[0]["借方金額(円)"] = CommaSeparate((res + Convert.ToInt32(dtresult.Rows[0]["貸方金額(円)"].ToString().Replace(",", ""))).ToString());
            dtresult.Columns.Remove("Rowno");
            dtresult.AcceptChanges();
            var Numberitems = Convert.ToInt32(PETC0308IBL.Mshiwake_title().Rows[0]["NumberOfItems"].ToString());
            if (dtresult.Rows.Count <= Numberitems)
            {
                DataTable folder = new DataTable();
                folder = PETC0308IBL.GetFolder();
                string path = string.Empty;
                path = folder.Rows[0]["Char2"].ToString();
                if (!string.IsNullOrWhiteSpace(path))
                {
                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }
                    //CsvWriter csvwriter = new CsvWriter();
                    //csvwriter = new CsvWriter();
                    //csvwriter.WriteCsv(dtresult, path + "\\" + filename, Encoding.GetEncoding(932));    -------------- Commented it bcos it include headers
                    //CSV(dtresult, path + "\\" + filename);
                    //DataTableToCSV.Program pj = new DataTableToCSV.Program();
                    //pj.CSV(dtresult, path + "\\" + filename, false);
                    ToCSV(dtresult, path + "\\" + filename);
                    Process.Start(Path.GetDirectoryName(path));
                }
            }
            else
            {
                var TotalSheet = NumberOfSheet(dtresult.Rows.Count, Numberitems);
                //SplittingDT.Chuncked chu = new Chuncked();
                //var Dt = chu.Separate(dtresult, Numberitems);
                SplitDT Sdt = new SplitDT();
                var Dt = Sdt.ListData(dtresult, Numberitems); 
                int hertz = 0;
                DataTable folder = new DataTable();
                folder = PETC0308IBL.GetFolder();
                string path = string.Empty;
                path = folder.Rows[0]["Char2"].ToString();
                foreach (var listdt in Dt)
                {

                    var cloned = listdt.Clone();
                    //       this.ImportedData = cloned;
                    cloned.Columns["仕訳メモ"].DataType = typeof(string);

                    foreach (DataRow dr in listdt.Rows)
                    {
                        cloned.ImportRow(dr);
                        //  cloned.lo;  
                    }


                    hertz = hertz + 1;
                    if (!string.IsNullOrWhiteSpace(path))
                    {
                        if (!Directory.Exists(path))
                        {
                            Directory.CreateDirectory(path);
                        }
                        foreach (DataRow dr in cloned.Rows)
                        {
                            dr["仕訳メモ"] = hertz + "/"+TotalSheet;
                        }
                        ToCSV(cloned, path + "\\" + filename.Split('.')[0].ToString() + "_" + hertz.ToString().PadLeft(2, '0')+".CSV");
                    }
                }
                Process.Start(Path.GetDirectoryName(path));
            }
        }
        public DataTable CommmaSeparatedinOutput(DataTable dt1, string[] cols)
        {
            DataTable dtCloned = dt1.Clone();   // Changed data type as string as fileds of sql selected column may be decimal or double type being of
            for (int i = 0; i < dtCloned.Columns.Count; i++)
            {
                if (dtCloned.Columns[i].ColumnName.Contains("借方金額(円)") || dtCloned.Columns[i].ColumnName.Contains("貸方金額(円)"))
                {
                    dtCloned.Columns[i].DataType = typeof(string);
                }

            }
            foreach (DataRow dr in dt1.Rows)
            {
                dtCloned.ImportRow(dr);
            }
            // var dt= new DataTable();
            foreach (string col in cols) 
            foreach (DataRow dr in dtCloned.Rows) { dr[col] = Convert.ToDecimal(dr[col].ToString()).ToString("#,##0"); }
            return dtCloned;

        }
        public static void ToCSV(DataTable dtDataTable, string strFilePath)
        {
            StreamWriter sw = new StreamWriter(strFilePath, false, Encoding.GetEncoding(932));
            //headers  
            //for (int i = 0; i < dtDataTable.Columns.Count; i++)
            //{
            //    sw.Write(dtDataTable.Columns[i]);
            //    if (i < dtDataTable.Columns.Count - 1)
            //    {
            //        sw.Write(",");
            //    }
            //}
            //sw.Write(sw.NewLine);
            foreach (DataColumn column in dtDataTable.Columns)
            {
                var value = String.Format("\"{0}\"", column.ColumnName);
                sw.Write(value);
                if (column.Ordinal < dtDataTable.Columns.Count)
                {
                    sw.Write(",");
                }
            }
            sw.Write(sw.NewLine);
            foreach (DataRow dr in dtDataTable.Rows)
            {
                for (int i = 0; i < dtDataTable.Columns.Count; i++)
                {
                    if (!Convert.IsDBNull(dr[i]))
                    {
                        string value = dr[i].ToString();
                        //if (value.Contains(','))
                        //{

                        value = String.Format("\"{0}\"", value);
                        sw.Write(value);

                        //}
                        //else
                        //{
                        //   sw.Write(dr[i].ToString());
                        //}
                    }
                    if (i < dtDataTable.Columns.Count - 1)
                    {
                        sw.Write(",");
                    }
                }
                sw.Write(sw.NewLine);
            }
            sw.Close();
        }
        private bool ErrorCheckForGrid()
        {
            if (string.IsNullOrWhiteSpace(txtImport.Text))
            {
                DSP_MSG("E102", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                txtImport.Focus();
                return false;
            }

            if (string.IsNullOrWhiteSpace(txtdocDate.Text))
            {
                DSP_MSG("E102", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                txtdocDate.Focus();
                return false;
            }

            int Row;
            int Column;
            var tupledate = Getlocation(dgvPETC0308I, "");
            Row = Convert.ToInt32(tupledate.Item1);
            Column = Convert.ToInt32(tupledate.Item2);
            if (Row.ToString() != "-1" && Column.ToString() != "-1")
            {
                DSP_MSG("E102", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);

                dgvPETC0308I.CurrentCell = dgvPETC0308I[Column, Row];
                return false;
            }
            var tupleminus = Getlocation(dgvPETC0308I, "-");
            Row = Convert.ToInt32(tupleminus.Item1);
            Column = Convert.ToInt32(tupleminus.Item2);
            if (Row.ToString() != "-1" && Column.ToString() != "-1")
            {
                DSP_MSG("E109", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                dgvPETC0308I.CurrentCell = dgvPETC0308I[Column, Row];
                return false;
            }
            return true;
        }

        public Tuple<string, string> Getlocation(DataGridView dg, string Check_params)
        {
            var colindex = "-1";
            var rowindex = "-1";
            foreach (DataGridViewRow row in dgvPETC0308I.Rows)
            {
                if (Check_params == "")
                {
                    if ((Convert.ToBoolean(row.Cells["coltesu_chk"].EditedFormattedValue) == true) && (row.Cells["coldocDate"].Value == null))
                    {
                        colindex = row.Cells["coldocDate"].ColumnIndex.ToString();
                        rowindex = row.Cells["coldocDate"].RowIndex.ToString();
                        goto Found;
                    }
                }
                else
                {
                    for (int i = 1; i <= 15; i++)
                    {
                        var value = (i).ToString().PadLeft(2, '0');
                        if (row.Cells["colKeihi" + value + "input"].Value.ToString().Contains(Check_params))
                        {
                            colindex = row.Cells["colKeihi" + value + "input"].ColumnIndex.ToString();
                            rowindex = row.Cells["colKeihi" + value + "input"].RowIndex.ToString();
                            //break;
                            goto Found;
                        }
                    }
                }

            }
        Found:
            return Tuple.Create(rowindex, colindex);
        }
        /// <summary>
        /// Button Click For F11
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnTorikomi_Click(object sender, EventArgs e)
        {
            F11();
        }
        protected void F11()
        {
            if (!string.IsNullOrWhiteSpace(txtImport.Text))
            {
                if (!string.IsNullOrWhiteSpace(txtdocDate.Text))
                {
                    if (DSP_MSG("Q101", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty).Equals("1"))
                    {
                        Cursor.Current = Cursors.WaitCursor;
                        filename = txtImport.Text;
                        Tessuryou_data = new T_Tesuuryou_Entity();
                        NyuukinMeisai_data = new T_NyuukinMeisai_Entity();
                        LLogEntity = new L_Log_Entity();
                        LLogEntity.InsertOperator = Tessuryou_data.OperatorCD = NyuukinMeisai_data.OperatorCD = loginInfo.OperatorCode;
                        LLogEntity.Program = "PETC0308I";
                        LLogEntity.PC = loginInfo.PcName;
                        LLogEntity.InsertDateTime = Tessuryou_data.InsertDateTime = NyuukinMeisai_data.InsertDatetime = System.DateTime.Now.ToString();
                        dtCSV = CSVToDatatable(filename);
                        if (datatableCheckColumn(dtCSV))
                        {
                            //Delete All data of T_NyuukinMeisai and T_Tesuuryou Table
                            if (PETC0308IBL.T_NyuukinMeisai_T_Tesuuryou_Delete(NyuukinMeisai_data, Tessuryou_data, LLogEntity))
                            {
                                // テーブル転送仕様Ａ ,テーブル転送仕様Ｂ
                                foreach (DataRow dr in dtCSV.Rows)
                                {
                                    dr["入金額"] = Convert.ToInt32(dr["入金額"].ToString().Split('.').First().ToString()).ToString();
                                }
                                try
                                {
                                    if ((dtCSV.Select("口座入金日 <> ''").CopyToDataTable().Rows.Count) > 0)
                                    {
                                        //if (dtCSV.Select("入金額 <> '0' AND ( 口座入金日 <> '' or 口座入金日 is not null)").CopyToDataTable().Rows.Count > 0)
                                        //{
                                        dtNyuukin = dtCSV.Select("入金額 <> '0' AND 口座入金日 <> '' ").CopyToDataTable();

                                        // SplittingDT.Chuncked ch = new Chuncked();
                                        //var f= ch.Separate(dtNyuukin,10000);
                                        //foreach (DataTable dt in f)
                                        //{
                                        NyuukinMeisai_data.xml = baseBL.DataTableToXml(dtNyuukin);

                                        // NyuukinMeisai_data.xml = baseBL.DataTableToXml(dtNyuukin);
                                        if (CheckMotoName(NyuukinMeisai_data))
                                        {
                                            if (PETC0308IBL.PETC0308I_Insert(NyuukinMeisai_data, Tessuryou_data, LLogEntity))
                                            {
                                                // 画面転送表01 and 画面転送表02
                                                BindGrid(txtdocDate.Text);
                                            }

                                            //}

                                        }
                                        //}

                                    }
                                }

                                catch
                                {

                                    DialogResult dialogResult = MessageBox.Show("There is no data inside fileds, 口座入金日 or 入金額 ! " + Environment.NewLine + " PLease set up the required Fileds.", "Information!", MessageBoxButtons.OK);
                                    if (dialogResult == DialogResult.OK)
                                    {
                                        //
                                    }
                                }
                            }

                        }
                    }
                }
                else
                {
                    DSP_MSG("E102", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                    txtdocDate.Focus();
                }
            }
            else
            {
                DSP_MSG("E102", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                txtImport.Focus();
            }
        }
        protected bool CheckMotoName(T_NyuukinMeisai_Entity NyuukinMeisai)
        {
            if (PETC0308IBL.Check_Motonames_MF(NyuukinMeisai))
            {
                DialogResult dialogResult = MessageBox.Show("Some Fields of '口座入金区分' Columns do not exist inside M_NyuukinMoto Table!. " + Environment.NewLine + " PLease set up the required Fileds. Do you want to proceed it by leaving extra columns anyway?", "Comfirmation!", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    return true;
                }
                else if (dialogResult == DialogResult.No)
                {
                    return false;
                }
            }
            return true;

        }
        private DataTable CSVToDatatable(string filePathName)
        {
            DataTable csvData = new DataTable();
            int count = 1;
            try
            {
                using (TextFieldParser csvReader = new TextFieldParser(filePathName, Encoding.GetEncoding(932), true))
                {
                    csvReader.SetDelimiters(new string[] { "," });
                    csvReader.HasFieldsEnclosedInQuotes = true;
                    //read column names
                    string[] colFields = csvReader.ReadFields();

                    foreach (string column in colFields)
                    {
                        if (firstColHeader.Equals("1")) //first row is column name
                        {
                            if (!csvData.Columns.Contains(column))
                            {
                                DataColumn datacolumn = new DataColumn(column);
                                datacolumn.AllowDBNull = true;
                                csvData.Columns.Add(datacolumn);
                            }
                            else
                            {
                                DataColumn datacolumn = new DataColumn(column + "_" + count++);
                                datacolumn.AllowDBNull = true;
                                csvData.Columns.Add(datacolumn);
                            }
                        }
                        else//2
                        {
                            DataColumn datacolumn = new DataColumn();
                            datacolumn.AllowDBNull = true;
                            csvData.Columns.Add(column);
                        }
                    }

                    if (firstColHeader.Equals("2")) // first row is data
                        csvData.Rows.Add(colFields);//add first row as data row

                    while (!csvReader.EndOfData)
                    {
                        string[] fieldData = csvReader.ReadFields();
                        while (fieldData.Count() > colFields.Count())
                        {
                            fieldData = fieldData.Take(fieldData.Count() - 1).ToArray();
                        }

                        while (fieldData.Count() < colFields.Count())
                        {
                            Array.Resize(ref fieldData, fieldData.Length + 1);
                            fieldData[fieldData.Length - 1] = null;
                        }

                        //Making empty value as null
                        for (int i = 0; i < fieldData.Length; i++)
                        {
                            if (fieldData[i] == "")
                            {
                                fieldData[i] = null;
                            }
                        }

                        for (int p = 0; p < fieldData.Count(); p++)
                        {
                            if (p == 6)
                            {
                                if (fieldData[p] != null)
                                {
                                    fieldData[p] = Convert.ToDateTime(fieldData[p]).ToString();
                                }
                            }
                        }
                        csvData.Rows.Add(fieldData);
                    }
                }
            }
            catch (Exception ex)
            {

            }
            return csvData;
        }

        private bool datatableCheckColumn(DataTable dt)
        {
            if (!(dt.Columns[0].ColumnName == "受注番号") && !(dt.Columns[1].ColumnName == "Web受注番号"))
            {
                DSP_MSG("E137", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                return false;
            }
            return true;
        }
        public void BindGrid(string date)
        {
            DataTable dtShow = new DataTable();
            DataTable dtDetail = new DataTable();
            Tessuryou_data = new T_Tesuuryou_Entity();
            NyuukinMeisai_data = new T_NyuukinMeisai_Entity();
            Tessuryou_data.OperatorCD = NyuukinMeisai_data.OperatorCD = loginInfo.OperatorCode;
            dtShow = PETC0308IBL.PETC0308I_SELECT_MF(Tessuryou_data);
            DataTable dtCloned = dtShow.Clone();   // Changed data type as string as fileds of sql selected column may be decimal or double type being of
            for (int i = 0; i < dtCloned.Columns.Count; i++)
            {
                if (dtCloned.Columns[i].ColumnName.Contains("input"))
                {
                    dtCloned.Columns[i].DataType = typeof(string);
                }

            }
            foreach (DataRow dr in dtShow.Rows)
            {
                dtCloned.ImportRow(dr);
            }
            dtCloned.Columns.Add("docDate");
            dtCloned.Columns.Add("tesu_chk", typeof(bool));
            foreach (DataRow row in dtCloned.Rows)
            {
                row["tesu_chk"] = true;
            }
            dgvPETC0308I.DataSource = dtCloned;



            for (int c = 1; c <= 15; c++)
            {
                decimal result = 0;
                foreach (DataGridViewRow gr in dgvPETC0308I.Rows)
                {
                    var val = (c).ToString().PadLeft(2, '0');
                    //   gr.Cells["colKeihi" + val + "Tax"].Value = Calculate_Tax(gr.Cells["colKeihi" + val + "TaxKBN"].Value.ToString(), gr.Cells["colTaxRate"].Value.ToString(), gr.Cells["colKeihi" + val + "input"].Value.ToString().Replace(",", ""));
                    gr.Cells["colKeihi" + val + "Tax"].Value = Calculate_Tax(gr.Cells["colKeihi" + val + "TaxKBN"].Value.ToString(), GetTaxtRate(gr.Cells["coldocDate"].EditedFormattedValue), gr.Cells["colKeihi" + val + "input"].Value.ToString().Replace(",", ""));
                    result = result + Convert.ToDecimal(gr.Cells["colKeihi" + val + "Input"].Value);
                    if (gr.Index + 1 == dgvPETC0308I.Rows.Count)
                    {
                        Control[] controls = this.Controls.Find("lbl_total" + (c), true);
                        if (controls.Length == 1) // 0 means not found, more - there are several controls with the same name
                        {
                            CKM_Controls.CKM_Label control = controls[0] as CKM_Controls.CKM_Label;
                            control.value = result.ToString();
                        }
                    }

                    if (String.IsNullOrEmpty(gr.Cells["colKeihi" + val + "Name"].Value.ToString()) || gr.Cells["colKeihi" + val + "Name"].Equals(DBNull.Value) || gr.Cells["colKeihi" + val + "Name"].FormattedValue.ToString() == "")
                    {
                        gr.Cells["colKeihi" + val + "input"].ReadOnly = true;
                    }
                    gr.Cells["colKeihi" + val + "input"].ToolTipText = gr.Cells["colKeihi" + val + "Name"].Value.ToString();
                }
            }

            foreach (DataGridViewRow gr in dgvPETC0308I.Rows)
            {
                gr.Cells["coldocDate"].Value = date;
                // gr.Cells["col_nyuukingaku"].Value = Convert.ToDecimal(gr.Cells["col_nyuukingaku"].Value).ToString("#,##0");
            }


            dtDetail = PETC0308IBL.PETC0308I_Detail_SELECT(NyuukinMeisai_data);
            dtDetail.Columns.Add("chk", typeof(bool));
            foreach (DataRow row in dtDetail.Rows)
            {
                row["chk"] = true;
            }
            dgvDetail.DataSource = dtDetail;

            foreach (DataGridViewRow dr in dgvDetail.Rows)
            {
                //dr.Cells["colNyuukinGaku"].Value = Convert.ToDecimal( dr.Cells["colNyuukinGaku"].Value).ToString("#,##0");
            }
            //CheckALL(dgvPETC0302I, "coltesu_chk",null,true,null);
            //CheckALL(dgvDetail, "colchk", null, true, null);

        }

        public static decimal Round_Off(string ipt)
        {
            decimal value;
            if (decimal.TryParse(ipt, out value))
            {
                //value = Math.Floor(value);
                value = Math.Ceiling(value);
                ipt = value.ToString();
                // Do something with the new text value
            }
            else
            {
                // Tell the user their input is invalid
            }
            return value;
        }

        public decimal Calculate_Tax(string kbn, string taxrate, string input)
        {
            input = string.IsNullOrWhiteSpace(input) ? "0" : input;   ///Before entering to dgvdadataerror event ,  Cos initial input is null or "" , it putting assign as "0"  
            decimal value;
            if (kbn == "1")
            {
                value = Convert.ToDecimal(input) * Convert.ToDecimal(taxrate) * Convert.ToDecimal("0.01");
            }
            else if (kbn == "2")
            {
                value = Convert.ToDecimal(input) * (Convert.ToDecimal(taxrate) / (Convert.ToDecimal(taxrate) + Convert.ToDecimal("100")));
            }
            else
            {
                value = 0;
            }

            return Round_Off(value.ToString());
        }

        private string GetTaxtRate(object date)
        {
            string d = String.IsNullOrWhiteSpace(Convert.ToString(date)) ? "0" : date.ToString();
            PETC0308IBL = new PETC0308I_BL();
            DataTable dt = new DataTable();

            if (d.Equals("0"))
            {
                return "0";
            }
            else
            {
                return PETC0308IBL.GetTaxRate(d).Rows[0]["TaxRate"].ToString();
            }
        }
        public void dgvPETC0302I_Sorted(object sender, EventArgs e)
        {

            try
            {
                BindGrid_Sorting();
                //dgvPETC0302I.Sort(dgvPETC0302I.Columns[(sender as DataGridView).SortedColumn.Name], ListSortDirection.Ascending);
                //dgvPETC0302I.Sort(dgvPETC0302I.Columns[(sender as DataGridView).SortedColumn.Name], ListSortDirection.Ascending);
            }
            catch { }


        }


        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {

            // Update the status bar when the cell value changes.

            //if (e.ColumnIndex == 0 && e.RowIndex != -1)
            //{

            //    // Get the orderID from the OrderID column.

            //    int orderID = (int)dgvDetail.Rows[e.RowIndex].Cells["colNo1"].Value;

            //    checkState[orderID] = (bool)dgvDetail.Rows[e.RowIndex].Cells[0].Value;

            //}
        }
        private void dataGridView1_CellValueNeeded(object sender, DataGridViewCellValueEventArgs e)
        {

            // Handle the notification that the value for a cell in the virtual column

            // is needed. Get the value from the dictionary if the key exists.



            //if (e.ColumnIndex == 0)
            //{

            //    int orderID = (int)dgvDetail.Rows[e.RowIndex].Cells["colNo1"].Value;

            //    if (checkState.ContainsKey(orderID))
            //    {

            //        e.Value = checkState[orderID];

            //    }

            //    else

            //        e.Value = false;

            //}



        }
        private void dataGridView1_CellValuePushed(object sender, DataGridViewCellValueEventArgs e)
        {

            // Handle the notification that the value for a cell in the virtual column

            // needs to be pushed back to the dictionary.



            //if (e.ColumnIndex == 0)
            //{

            //    // Get the orderID from the OrderID column.

            //    int orderID = (int)dgvDetail.Rows[e.RowIndex].Cells["colNo1"].Value;



            //    // Add or update the checked value to the dictionary depending on if the

            //    // key (orderID) already exists.

            //    if (!checkState.ContainsKey(orderID))
            //    {

            //        checkState.Add(orderID, (bool)e.Value);

            //    }

            //    else

            //        checkState[orderID] = (bool)e.Value;

            //}

        }
        private System.Collections.Generic.Dictionary<int, bool> checkState;
        public void BindGrid_Sorting()
        {
            for (int c = 1; c <= 15; c++)
            {
                decimal result = 0;
                foreach (DataGridViewRow gr in dgvPETC0308I.Rows)
                {
                    var val = (c).ToString().PadLeft(2, '0');
                    //   gr.Cells["colKeihi" + val + "Tax"].Value = Calculate_Tax(gr.Cells["colKeihi" + val + "TaxKBN"].Value.ToString(), gr.Cells["colTaxRate"].Value.ToString(), gr.Cells["colKeihi" + val + "input"].Value.ToString().Replace(",", ""));
                    gr.Cells["colKeihi" + val + "Tax"].Value = Calculate_Tax(gr.Cells["colKeihi" + val + "TaxKBN"].Value.ToString(), GetTaxtRate(gr.Cells["coldocDate"].EditedFormattedValue), gr.Cells["colKeihi" + val + "input"].Value.ToString().Replace(",", ""));
                    result = result + Convert.ToDecimal(gr.Cells["colKeihi" + val + "Input"].Value);
                    if (gr.Index + 1 == dgvPETC0308I.Rows.Count)
                    {
                        Control[] controls = this.Controls.Find("lbl_total" + (c), true);
                        if (controls.Length == 1) // 0 means not found, more - there are several controls with the same name
                        {
                            CKM_Controls.CKM_Label control = controls[0] as CKM_Controls.CKM_Label;
                            control.value = result.ToString();
                        }
                    }

                    if (String.IsNullOrEmpty(gr.Cells["colKeihi" + val + "Name"].Value.ToString()) || gr.Cells["colKeihi" + val + "Name"].Equals(DBNull.Value) || gr.Cells["colKeihi" + val + "Name"].FormattedValue.ToString() == "")
                    {
                        gr.Cells["colKeihi" + val + "input"].ReadOnly = true;
                    }
                    gr.Cells["colKeihi" + val + "input"].ToolTipText = gr.Cells["colKeihi" + val + "Name"].Value.ToString();
                }
            }

            //foreach (DataGridViewRow gr in dgvPETC0302I.Rows)
            //{
            //    gr.Cells["col_nyuukingaku"].Value = Convert.ToDecimal(gr.Cells["col_nyuukingaku"].Value).ToString("#,##0");
            //}
            //foreach (DataGridViewRow dr in dgvDetail.Rows)
            //{
            //    dr.Cells["colNyuukinGaku"].Value = Convert.ToDecimal(dr.Cells["colNyuukinGaku"].Value).ToString("#,##0");
            //}

        }

        private void txtdocDate_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (!string.IsNullOrWhiteSpace(txtdocDate.Text))
                {
                    if ((txtdocDate.Text).Contains("/"))
                    {
                        foreach (DataGridViewRow gr in dgvPETC0308I.Rows)
                        {
                            gr.Cells["coldocDate"].Value = txtdocDate.Text;
                        }
                    }
                }
                else
                    txtdocDate.Focus();
            }
        }

        private void dgvPETC0308I_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            
            dgcell = (sender as DataGridView).CurrentCell;   // store this value when reselection on error dgv occurred (may not be used)
       
        }

        private void dgvPETC0308I_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex != -1)   // Error skiiped when not clicking on header bar
            {
                if (dgvPETC0308I.CurrentCell.ColumnIndex != 0 && dgvPETC0308I.CurrentCell.ColumnIndex != 3 && dgvPETC0308I.CurrentCell.ColumnIndex != 1 && e.RowIndex >= 0)  // must not be these rows
                {
                    dgvPETC0308I.CurrentCell = dgvPETC0308I.Rows[e.RowIndex].Cells[e.ColumnIndex];
                    dgvPETC0308I.CurrentCell.Selected = true;  /// as soon as cursor point set the current cells, making it the cell edit
                    dgvPETC0308I.BeginEdit(true);  // make event for setting flags
                }
            }
        }

        private void dgvPETC0308I_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            Maintained_CellClick(sender, e);
        }

        private void dgvPETC0308I_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            Maintained_CellClick(sender, e);
        }
        protected void Maintained_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if ((sender as DataGridView).CurrentCell is DataGridViewCheckBoxCell)
            {
                string MotoCD = "0";
                string KouzaDate = "0";
                int res = 0;
                if (e.RowIndex != -1)
                {
                    if ((Convert.ToBoolean(dgvPETC0308I.Rows[e.RowIndex].Cells[e.ColumnIndex].EditedFormattedValue) == true))
                    {
                        MotoCD = dgvPETC0308I.Rows[e.RowIndex].Cells["colNyuukinMotoCD"].Value.ToString();
                        KouzaDate = dgvPETC0308I.Rows[e.RowIndex].Cells["colKouzaNyuukinDate"].Value.ToString();
                        string[] moto_kouza = { MotoCD, KouzaDate };
                        CheckALL((sender as DataGridView), "blur", moto_kouza, true);

                        foreach (DataGridViewRow row1 in dgvDetail.Rows)
                        {
                            if (row1.Cells["colNyuukinMotoCD1"].Value.ToString() == MotoCD && row1.Cells["colKouzaNyuukinDate1"].Value.ToString() == KouzaDate)
                            {
                                res += Convert.ToInt32(row1.Cells["colNyuukinGaku"].Value.ToString().Replace(",", string.Empty));
                                row1.Cells["colchk"].Value = true;
                            }
                        }
                    }
                    else
                    {
                        MotoCD = dgvPETC0308I.Rows[e.RowIndex].Cells["colNyuukinMotoCD"].Value.ToString();
                        KouzaDate = dgvPETC0308I.Rows[e.RowIndex].Cells["colKouzaNyuukinDate"].Value.ToString();
                        string[] moto_kouza = { MotoCD, KouzaDate };
                        CheckALL((sender as DataGridView), "blur", moto_kouza, false);
                        foreach (DataGridViewRow row1 in dgvDetail.Rows)
                        {
                            if (row1.Cells["colNyuukinMotoCD1"].Value.ToString() == MotoCD && row1.Cells["colKouzaNyuukinDate1"].Value.ToString() == KouzaDate)
                            {

                                row1.Cells["colchk"].Value = false;
                            }
                        }
                    }
                }

                foreach (DataGridViewRow row1 in dgvPETC0308I.Rows)
                {
                    var d = row1.Index;
                    if (row1.Cells["colNyuukinMotoCD"].Value.ToString() == MotoCD && row1.Cells["colKouzaNyuukinDate"].Value.ToString() == KouzaDate)
                    {
                        row1.Cells["col_nyuukingaku"].Value = res.ToString();
                    //   row1.Cells["col_nyuukingaku"].Value = CommaSeparate(res.ToString());
                    }
                }
            }
        }

        private void dgvPETC0308I_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex >= (sender as DataGridView).Columns["colKeihi01input"].Index)
            {
                Sum_hiddenFiled(e.ColumnIndex);       //Enforeced ... make summation when ending in edit mode as a recover if someoneexceptions occurred
                dgvPETC0308I.NotifyCurrentCellDirty(true);   // put this not occuring win box errors for abnormal changing in cell values
            }
        }
        public void Sum_hiddenFiled(int colindex)  // make summation from all taxes calculated stored in hidded fields
        {
            decimal result = 0;
            var col_num = dgvPETC0308I.Columns[colindex].DataPropertyName.ToString().Replace("Keihi", "").Replace("Input", "").Replace("input", "");
            foreach (DataGridViewRow gr in dgvPETC0308I.Rows)
            {
                result = result + Convert.ToDecimal(gr.Cells["colKeihi" + col_num + "Input"].Value);
                if (gr.Index + 1 == dgvPETC0308I.Rows.Count)
                {
                    Control[] controls = this.Controls.Find("lbl_total" + (Convert.ToInt32(col_num).ToString()), true);
                    if (controls.Length == 1) // 0 means not found, more - there are several controls with the same name
                    {
                        CKM_Controls.CKM_Label control = controls[0] as CKM_Controls.CKM_Label;
                        control.value = result.ToString("#,##0");
                    }
                }
            }
        }

        /// <summary>
        /// Check The Input Value for Date is Integer
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        private bool IsInteger(string value)
        {
            int Num;
            if (Int32.TryParse(value, out Num))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Check The Input Value for Date is Correct date
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        private bool CheckDate(string value)
        {
            DateTime d;
            var g = DateTime.TryParseExact(value, "yyyy/MM/dd", System.Globalization.CultureInfo.InvariantCulture, DateTimeStyles.None, out d);
            return g;
        }
        /// <summary>
        /// Check The Input Date is Correct and Convert It To Date Format
        /// </summary>
        /// <param name="docDate"></param>
        /// <returns></returns>
        /// 
        private string checkDate(string docDate, DataGridViewCell dcell)
        {
            bool flg = true;
            if (docDate.Contains("/"))    //Condition for 2019/09/9  >>>to be true  by ptk
            {
                string day = string.Empty; ;
                if (docDate.Split('/').Count() == 3 && docDate.Split('/').First().Length == 4 && docDate.Split('/')[1].Length == 2)
                {
                    day = docDate.Split('/')[2].PadLeft(2, '0');
                    if (CheckDate(docDate.Split('/')[0] + "/" + docDate.Split('/')[1] + "/" + day))
                    {
                        return docDate;
                    }
                    else
                    {
                        DSP_MSG("E103", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                        flg = false;
                    }
                }
                else
                {
                    DSP_MSG("E103", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                    flg = false;
                }
            }
            else
            {
                if (IsInteger(docDate.Replace("/", "")))
                {
                    string day = string.Empty, month = string.Empty, year = string.Empty;
                    if (docDate.Contains("/"))
                    {
                        string[] date = docDate.Split('/');
                        day = date[date.Length - 1].PadLeft(2, '0');
                        month = date[date.Length - 2].PadLeft(2, '0');

                        if (date.Length > 2)
                            year = date[date.Length - 3];

                        docDate = docDate.Replace("/", "");
                    }

                    docDate = docDate.PadLeft(8, '0');
                    day = docDate.Substring(docDate.Length - 2);
                    month = docDate.Substring(docDate.Length - 4).Substring(0, 2);
                    year = Convert.ToInt32(docDate.Substring(0, docDate.Length - 4)).ToString();

                    if (month == "00")
                    {
                        month = string.Empty;
                    }
                    if (year == "0")
                    {
                        year = string.Empty;
                    }

                    if (string.IsNullOrWhiteSpace(month))
                        month = DateTime.Now.Month.ToString().PadLeft(2, '0');//if user doesn't input for month,set current month

                    if (string.IsNullOrWhiteSpace(year))
                    {
                        year = DateTime.Now.Year.ToString();//if user doesn't input for year,set current year
                    }
                    else
                    {
                        if (year.Length == 1)
                            year = "200" + year;
                        else if (year.Length == 2)
                            year = "20" + year;
                    }

                    string strdate = year + "/" + month + "/" + day;
                    if (CheckDate(strdate))
                    {
                        //dgv1.Rows[e.RowIndex].Cells[4].Value = strdate;
                        return strdate;
                    }
                    else
                    {
                        DSP_MSG("E103", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                        flg = false;
                    }
                }
            }

            if (flg)
            {
                //Make 
            }
            return null;

        }
        private void dgvPETC0308I_CellValidated(object sender, DataGridViewCellEventArgs e)
        {

            if ((sender as DataGridView).CurrentCell is DataGridViewTextBoxCell)
            {
                string chk_date = null;
                string[] dn = null;
                if ((Convert.ToBoolean(dgvPETC0308I.Rows[e.RowIndex].Cells["coltesu_chk"].EditedFormattedValue) == true) && (dgvPETC0308I.CurrentCell == dgvPETC0308I.Rows[e.RowIndex].Cells["coldocDate"]))
                {
                    if (dgvPETC0308I.Rows[e.RowIndex].Cells["coldocDate"].Value == null)
                    {
                        DSP_MSG("E102", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                    }
                    else
                    {
                        Date = dgvPETC0308I.Rows[e.RowIndex].Cells["coldocDate"].Value.ToString();
                        chk_date = checkDate(Date, dgvPETC0308I.CurrentCell);
                        dgvPETC0308I.Rows[e.RowIndex].Cells["coldocDate"].Value = chk_date;
                        if (chk_date != null)
                        {
                            MotoCD = dgvPETC0308I.Rows[e.RowIndex].Cells["colNyuukinMotoCD"].Value.ToString();
                            KouzaDate = dgvPETC0308I.Rows[e.RowIndex].Cells["colKouzaNyuukinDate"].Value.ToString();
                            string[] moto_kouza = { MotoCD, KouzaDate };
                            CheckALL((sender as DataGridView), "blur", moto_kouza, false, chk_date);
                        }
                    }
                }
                else if ((dgvPETC0308I.CurrentCell == dgvPETC0308I.Rows[e.RowIndex].Cells["coldocDate"]) && (!(dgvPETC0308I.Rows[e.RowIndex].Cells["coldocDate"].Value == null)))
                {
                    Date = dgvPETC0308I.Rows[e.RowIndex].Cells["coldocDate"].Value.ToString();
                    chk_date = checkDate(Date, dgvPETC0308I.CurrentCell);
                    dgvPETC0308I.Rows[e.RowIndex].Cells["coldocDate"].Value = chk_date;
                    if (chk_date != null)
                    {
                        MotoCD = dgvPETC0308I.Rows[e.RowIndex].Cells["colNyuukinMotoCD"].Value.ToString();
                        KouzaDate = dgvPETC0308I.Rows[e.RowIndex].Cells["colKouzaNyuukinDate"].Value.ToString();
                        string[] moto_kouza = { MotoCD, KouzaDate };
                        CheckALL((sender as DataGridView), "blur", moto_kouza, false, chk_date);
                    }
                }
            }
            if (((sender as DataGridView).CurrentCell.ColumnIndex >= (sender as DataGridView).CurrentCell.DataGridView.Columns["colKeihi01input"].Index))
            {
                if ((sender as DataGridView).CurrentCell.Value.ToString().Contains('-'))  // Check Minus
                {
                    DSP_MSG("E109", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                    set_focus = true;
                }
                else  // set value tax and set as separated comma values
                {
                    string input = (sender as DataGridView).CurrentCell.EditedFormattedValue.ToString().Replace(",", "");
                    //var tax_rate = dgvPETC0302I[(sender as DataGridView).CurrentCell.DataGridView.Columns["colTaxRate"].Index, (sender as DataGridView).CurrentCell.RowIndex].Value;

                    var tax_rate = dgvPETC0308I[(sender as DataGridView).CurrentCell.DataGridView.Columns["coldocDate"].Index, (sender as DataGridView).CurrentCell.RowIndex].Value;
                    var kbn_col = "colKeihi" + (sender as DataGridView).Columns[(sender as DataGridView).CurrentCell.ColumnIndex].DataPropertyName.ToString().Replace("Keihi", "").Replace("Input", "").Replace("input", "") + "TaxKBN";
                    var kbn = dgvPETC0308I.Rows[(sender as DataGridView).CurrentCell.RowIndex].Cells[kbn_col].Value;
                    var tax_col = "colKeihi" + (sender as DataGridView).Columns[(sender as DataGridView).CurrentCell.ColumnIndex].DataPropertyName.ToString().Replace("Keihi", "").Replace("Input", "") + "Tax";
                    (sender as DataGridView)[tax_col, (sender as DataGridView).CurrentCell.RowIndex].Value = Calculate_Tax(kbn.ToString(), GetTaxtRate(dgvPETC0308I.Rows[e.RowIndex].Cells["coldocDate"].EditedFormattedValue), input.ToString());
                    (sender as DataGridView).CurrentCell.Value = CommaSeparate(input);

                }
            }
        }

        private void dgvPETC0308I_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            (sender as DataGridView).EditingControl.ResetText();
            (sender as DataGridView).EditingControl.Text = "0";
            dgvPETC0308I.CellEndEdit += dgvPETC0308I_CellEndEdit;
            if (!TabSkip)
            {
                SendKeys.Send("{Tab}");    // It Needs when entering to dgvdataerror occurred by setting value "0"
                TabSkip = true;
            }
            else
            {
                TabSkip = false;
            }
            e.Cancel = true;
        }

        private void dgvPETC0308I_Sorted(object sender, EventArgs e)
        {

            try
            {
                BindGrid_Sorting();
                //dgvPETC0302I.Sort(dgvPETC0302I.Columns[(sender as DataGridView).SortedColumn.Name], ListSortDirection.Ascending);
                //dgvPETC0302I.Sort(dgvPETC0302I.Columns[(sender as DataGridView).SortedColumn.Name], ListSortDirection.Ascending);
            }
            catch { }
        }

        private void btnTorikomi_Click_1(object sender, EventArgs e)
        {
            F11();

        }
    }
}
