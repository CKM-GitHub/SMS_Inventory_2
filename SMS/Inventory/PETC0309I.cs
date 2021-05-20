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
    public partial class frmPETC0309I : SMS_Base
    {
        Login_Info loginInfo;
        PETC0309I_BL PETC0309IBL;
    //    PETC0309I_BL PETC0309IBL;
        Base_BL baseBL = new Base_BL();
        L_Log_Entity LLogEntity;
        string filename = string.Empty;
        string Date = string.Empty;
        string xml = string.Empty;
        string MotoCD = string.Empty;
        string KouzaDate = string.Empty;
        bool TabSkip = false;
        String[] files;
        T_Tesuuryou_Entity Tessuryou_data;
        T_NyuukinMeisai_Entity NyuukinMeisai_data;
        DataGridViewCell dgcell = null;
        public frmPETC0309I()
        {
            InitializeComponent();
        }
        public frmPETC0309I(Login_Info loginInfo)
            : base(loginInfo)
        {
            InitializeComponent();
            this.loginInfo = loginInfo;
        }

        private void frmPETC0309I_Load(object sender, EventArgs e)
        {

            lblMode.Visible = false;
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
            FormName = "入金手数料(店舗POS)仕訳作成MF";
            PETC0303IDetail.EnabledColumn("colchk");
            dgvPETC0303I.Columns["colNo"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            PETC0303IDetail.Columns["colNo1"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgvPETC0303I.DisabledColumn("colNo,colimpName,colKouzaNyuukinDate,col_nyuukingaku");
            PETC0303IDetail.Gridcolor(Color.Black);
            dgvPETC0303I.Gridcolor(Color.Black);
            PETC0309IBL = new PETC0309I_BL();
            txtdocDate.Text = DateTime.Now.ToString("yyyy/MM/dd").Replace("-","/");
            var dt = PETC0309IBL.Mshiwake_title();
            for (int j = 0; j < 15; j++)
            {
                var col = (j + 1).ToString().PadLeft(2, '0');
                dgvPETC0303I.Columns["colKeihi" + col + "Input"].ValueType = typeof(string);
                dgvPETC0303I.Columns["colKeihi" + col + "Input"].HeaderText = dt.Rows[0]["KeihiTitle" + col].ToString();
                Control[] controls = this.Controls.Find("lbl_total" + (j + 1), true);
                if (controls.Length == 1) // 0 means not found, more - there are several controls with the same name
                {
                    CKM_Controls.CKM_Label control = controls[0] as CKM_Controls.CKM_Label;
                    control.value = "0";
                }
            }
            DisableHeaderFlexible(dgvPETC0303I);
            DisableHeaderFlexible(PETC0303IDetail);
        }
        public void DisableHeaderFlexible(DataGridView dg)
        {
            dg.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;

            for (int i = 0; i < dg.Columns.Count; i++)
            {
                dg.Columns[i].Resizable = DataGridViewTriState.False;
                dg.Columns[i].Resizable = DataGridViewTriState.False;
            }
            //foreach (DataGridViewColumn column in dg.Columns)
            //{
            //    column.SortMode = DataGridViewColumnSortMode.NotSortable;
            //}
        }
        private string GetTaxtRate(object date)
        {
            string d = String.IsNullOrWhiteSpace(Convert.ToString(date)) ? "0" : date.ToString();
            PETC0309IBL = new PETC0309I_BL();
            DataTable dt = new DataTable();

            if (d.Equals("0"))
            {
                return "0";
            }
            else
            {
                return PETC0309IBL.GetTaxRate(d).Rows[0]["TaxRate"].ToString();
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
        private void F11()
        {
            if (!string.IsNullOrWhiteSpace(txtImport.Text))
            {
                if (!string.IsNullOrWhiteSpace(txtdocDate.Text))
                {
                    if (DSP_MSG("Q101", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty).Equals("1"))
                    {
                        Cursor.Current = Cursors.WaitCursor;
                        PETC0309IBL = new PETC0309I_BL();
                        Tessuryou_data = new T_Tesuuryou_Entity();
                        NyuukinMeisai_data = new T_NyuukinMeisai_Entity();
                        LLogEntity = new L_Log_Entity();
                        LLogEntity.InsertOperator = Tessuryou_data.OperatorCD = NyuukinMeisai_data.OperatorCD = loginInfo.OperatorCode;
                        LLogEntity.Program = "PETC0309I";
                        LLogEntity.PC = loginInfo.PcName;
                        LLogEntity.InsertDateTime = Tessuryou_data.InsertDateTime = NyuukinMeisai_data.InsertDatetime = System.DateTime.Now.ToString();
                        PopUp.PETC0309I_Import imp = new PopUp.PETC0309I_Import(files, LLogEntity);
                        imp.ShowDialog();

                        // 画面転送表01 and 画面転送表02
                        BindGrid(txtdocDate.Text);


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
        public void BindGrid(string date)
        {
            DataTable dtShow = new DataTable();
            DataTable dtDetail = new DataTable();
            PETC0309IBL = new PETC0309I_BL();
            DataSet ds = new DataSet();
            ds = PETC0309IBL.PETC0309I_SELECT(loginInfo.OperatorCode);
            dtShow = ds.Tables[0];

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

            dgvPETC0303I.DataSource = dtCloned;

            for (int c = 1; c <= 15; c++)
            {
                decimal result = 0;
                foreach (DataGridViewRow gr in dgvPETC0303I.Rows)
                {
                    var val = (c).ToString().PadLeft(2, '0');
                    //   gr.Cells["colKeihi" + val + "Tax"].Value = Calculate_Tax(gr.Cells["colKeihi" + val + "TaxKBN"].Value.ToString(), gr.Cells["colTaxRate"].Value.ToString(), gr.Cells["colKeihi" + val + "input"].Value.ToString().Replace(",", ""));
                    gr.Cells["colKeihi" + val + "Tax"].Value = Calculate_Tax(gr.Cells["colKeihi" + val + "TaxKBN"].Value.ToString(), GetTaxtRate(gr.Cells["coldocDate"].EditedFormattedValue), gr.Cells["colKeihi" + val + "input"].Value.ToString().Replace(",", ""));
                    result = result + Convert.ToDecimal(gr.Cells["colKeihi" + val + "Input"].Value);
                    if (gr.Index + 1 == dgvPETC0303I.Rows.Count)
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

            foreach (DataGridViewRow gr in dgvPETC0303I.Rows)
            {
                gr.Cells["coldocDate"].Value = date;
                //gr.Cells["col_nyuukingaku"].Value = Convert.ToDecimal(gr.Cells["col_nyuukingaku"].Value).ToString("#,##0");
            }

            dtDetail = ds.Tables[1];
            dtDetail.Columns.Add("chk", typeof(bool));
            foreach (DataRow row in dtDetail.Rows)
            {
                row["chk"] = true;
            }
            PETC0303IDetail.DataSource = dtDetail;

            //Commaprice_Column(PETC0303IDetail, new string[] {"coltransferAmount", "colcash", "colcheck", "colbill", "coloffset", "coladjustment" });
            // Commaprice_Column(PETC0303IDetail, new string[] { "coltransferAmount" });
            //CheckALL(dgvPETC0303I, "coltesu_chk", null, true, null);

            //CheckALL(PETC0303IDetail, "colchk", null, true, null);

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


        public static decimal Round_Off(string ipt)
        {
            //var fvalue = Math.Ceiling(Convert.ToDecimal(5.5));
            //var e33value = Math.Ceiling(Convert.ToDecimal(5.9));
            //var fvalusae = Math.Ceiling(Convert.ToDecimal(5.01));
            decimal value;
            if (decimal.TryParse(ipt, out value))
            {
                value = Math.Ceiling(value);
                ipt = value.ToString();
                //value = Math.Floor(value);
                //ipt = value.ToString();
                // Do something with the new text value
            }
            else
            {
                // Tell the user their input is invalid
            }
            return value;
        }


        private void F12()
        {
            if (ErrorCheckForGrid())
            {
                if (DSP_MSG("Q203", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty).Equals("1"))
                {
                    //テーブル転送仕様Ｃ, テーブル転送仕様Ｄ
                    if (PETC0309IBL.Update_Tesuyou_Nyuukin(Tesuyou_dt(), Nyuukin_dt()))  // Done
                    {
                        //出力転送表01, 出力転送表02
                        PETC0309IBL = new PETC0309I_BL();
                        DataTable dtexp = new DataTable();
                        NyuukinMeisai_data = new T_NyuukinMeisai_Entity();
                        NyuukinMeisai_data.OperatorCD = loginInfo.OperatorCode;
                        Export_CSV(PETC0309IBL.NyuukinDate_NyuukinMoto_Select(NyuukinMeisai_data.OperatorCD, "2"), PETC0309IBL.PETC0302I_CSVExport1(NyuukinMeisai_data, "PETC0309I_ExportCSV"));
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
        public void output_csv(DataTable dtout, string filename)
        {
            //if (dtout.Rows.Count == 1)
            //{
            //    dtout.Rows[0]["識別フラグ"] = "2111";
            //}
            //else
            //{
            //    dtout.Rows[0]["識別フラグ"] = "2110";
            //    dtout.Rows[dtout.Rows.Count - 1]["識別フラグ"] = "2101";
            //}


            DataTable dtresult = CommmaSeparatedinOutput(Cover_KeijouDate(dtout), new string[] { "借方金額(円)", "貸方金額(円)" });

            dtresult.Columns.Remove("NyuukinMotoCD");


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
                        if (i != 0)
                        {

                            // dtresult.Rows[i]["借方金額(円)"] ="0";
                            dtresult.Rows[i]["借方税額"] = "0";
                            // res += Convert.ToInt32(dtresult.Rows[i]["貸方金額(円)"].ToString().Replace(",", ""));
                        }
                        dtresult.Rows[i]["貸方税額"] = "0";
                      //  dtresult.Rows[i]["借方勘定科目"] = dtresult.Rows[i]["借方部門"] = dtresult.Rows[i]["借方補助科目"] = dtresult.Rows[i]["借方税区分"] = "";
                      ////  dtresult.Rows[i]["借方金額(円)"] = 
                      //  dtresult.Rows[i]["借方税額"] ="0";
                      // // res += Convert.ToInt32(dtresult.Rows[i]["貸方金額(円)"].ToString().Replace(",", ""));
                    }
                    dtresult.Rows[i]["貸方税額"] = "0";
                }
            }
         //   dtresult.Rows[0]["借方金額(円)"] = CommaSeparate((res + Convert.ToInt32(dtresult.Rows[0]["貸方金額(円)"].ToString().Replace(",", ""))).ToString());
            dtresult.Rows[0]["借方税額"] = "0";


            dtresult.Columns.Remove("Rowno");
            dtresult.Columns.Remove("KouZadate");
            dtresult.AcceptChanges();


            var Numberitems = Convert.ToInt32(PETC0309IBL.Mshiwake_title().Rows[0]["NumberOfItems"].ToString());
            if (dtresult.Rows.Count <= Numberitems)
            {
                DataTable folder = new DataTable();
                folder = PETC0309IBL.GetFolder3i();
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
                folder = PETC0309IBL.GetFolder3i();
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
                         dr["仕訳メモ"] = hertz + "/" + TotalSheet;
                        }
                        ToCSV(cloned, path + "\\" + filename.Split('.')[0].ToString() + "_" + hertz.ToString().PadLeft(2, '0') + ".CSV");
                    }
                }
                Process.Start(Path.GetDirectoryName(path));
            }
            //DataTable folder = new DataTable();
            //folder = PETC0309IBL.GetFolder3i();
            //string path = string.Empty;
            //path = folder.Rows[0]["Char2"].ToString();
            //if (!string.IsNullOrWhiteSpace(path))
            //{
            //    if (!Directory.Exists(path))
            //    {
            //        Directory.CreateDirectory(path);
            //    }

            //    //1st method
            //    //CsvWriter csvwriter = new CsvWriter();
            //    //csvwriter = new CsvWriter();
            //    //csvwriter.WriteCsv(dtresult, path + "\\" + filename, Encoding.GetEncoding(932)); //   -------------- Commented it bcos it include headers
            //    //2nd method
            //    //DataTableToCSV.Program pj = new DataTableToCSV.Program();
            //    //pj.CSV(dtresult, path + "\\" + filename, false);
            //    //3rd method
            //    //CSV(dtresult, path + "\\" + filename);
            //    //4th method
            //    ToCSV(dtresult, path + "\\" + filename);
            //    Process.Start(Path.GetDirectoryName(path + "\\"));
            //}
        }
        protected int NumberOfSheet(int totalCount, int NumberofItems)
        {
            if (totalCount == 0 || NumberofItems == 0)
            {
                return 0;
            }
            else
            {
                int res = 0;
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
                            drow["取引日付"] = string.IsNullOrEmpty(dr["KeijouDate"].ToString()) ? "" : dr["KeijouDate"].ToString();
                        }
                    }
                }
            }

            return dt;
        }
        public DataTable CommmaSeparatedinOutput(DataTable dt1, string[] cols)
        {


            DataTable dtCloned = dt1.Clone();   // Changed data type as string as fileds of sql selected column may be decimal or double type being of
            for (int i = 0; i < dtCloned.Columns.Count; i++) dtCloned.Columns[i].DataType = typeof(string);
            foreach (DataRow dr in dt1.Rows)
                dtCloned.ImportRow(dr);
            foreach (string col in cols) foreach (DataRow dr in dtCloned.Rows) dr[col] = Convert.ToDecimal(dr[col].ToString()).ToString("#,##0");
            //  "識別フラグ"
            var f = new string[3];
            f[0] = cols[0];
            f[1] = cols[1];
            // f[2] = "識別フラグ";
            foreach (string c in f)
                for (int a = 0; a < dtCloned.Columns.Count; a++)
                    foreach (DataRow dr in dtCloned.Rows)
                        if (dtCloned.Columns[a].ColumnName != c) dr[a] = dr[a].ToString().Replace("'", "'");
            return dtCloned;
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
            var tupledate = Getlocation(dgvPETC0303I, "");
            Row = Convert.ToInt32(tupledate.Item1);
            Column = Convert.ToInt32(tupledate.Item2);
            if (Row.ToString() != "-1" && Column.ToString() != "-1")
            {
                DSP_MSG("E102", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);

                dgvPETC0303I.CurrentCell = dgvPETC0303I[Column, Row];
                return false;
            }
            var tupleminus = Getlocation(dgvPETC0303I, "-");
            Row = Convert.ToInt32(tupleminus.Item1);
            Column = Convert.ToInt32(tupleminus.Item2);
            if (Row.ToString() != "-1" && Column.ToString() != "-1")
            {
                DSP_MSG("E109", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                dgvPETC0303I.CurrentCell = dgvPETC0303I[Column, Row];
                return false;
            }
            return true;
        }
        public Tuple<string, string> Getlocation(DataGridView dg, string Check_params)
        {
            var colindex = "-1";
            var rowindex = "-1";
            foreach (DataGridViewRow row in dgvPETC0303I.Rows)
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
            for (int j = 0; j < dgvPETC0303I.Rows.Count; j++)
            {
                dt.Rows.Add();
                dt.Rows[j]["OperatorCD"] = loginInfo.OperatorCode;
                dt.Rows[j]["DetaKBN"] = "2";
                dt.Rows[j]["coltesu_chk"] = Convert.ToBoolean(dgvPETC0303I.Rows[j].Cells["coltesu_chk"].EditedFormattedValue) ? "0" : "1";
                dt.Rows[j]["KeijouDate"] = dgvPETC0303I.Rows[j].Cells["coldocDate"].Value;
                dt.Rows[j]["NyuukinMotoCD"] = dgvPETC0303I.Rows[j].Cells["colNyuukinMotoCD"].Value;
                dt.Rows[j]["KouzaNyuukinDate"] = dgvPETC0303I.Rows[j].Cells["colKouzaNyuukinDate"].Value;
                dt.Rows[j]["SEQ"] = dgvPETC0303I.Rows[j].Cells["colseq"].Value;
            }

            for (int c = 1; c <= 15; c++)
            {
                foreach (DataGridViewRow gr in dgvPETC0303I.Rows)
                {
                    var val = (c).ToString().PadLeft(2, '0');
                    dt.Rows[gr.Index]["Keihi" + val + "Tax"] = gr.Cells["colKeihi" + val + "Tax"].Value;
                    dt.Rows[gr.Index]["Keihi" + val + "input"] = gr.Cells["colKeihi" + val + "input"].Value;
                }
            }
            return dt;
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

            foreach (DataGridViewRow gr in PETC0303IDetail.Rows)
            {
                dt.Rows.Add();
                dt.Rows[gr.Index]["OperatorCD"] = loginInfo.OperatorCode;
                dt.Rows[gr.Index]["colchk"] = Convert.ToBoolean(gr.Cells["colchk"].EditedFormattedValue) ? "0" : "1";
                dt.Rows[gr.Index]["DetaKBN"] = "2";
                dt.Rows[gr.Index]["colNyuukinMotoCD1"] = gr.Cells["colNyuukinMotoCD1"].Value;
                dt.Rows[gr.Index]["colKouzaNyuukinDate1"] = gr.Cells["colKouzaNyuukinDate1"].Value;
                dt.Rows[gr.Index]["colDetailNo"] = gr.Cells["colDetailNo"].Value;
            }
            return dt;
        }

        private void txtdocDate_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (!string.IsNullOrWhiteSpace(txtdocDate.Text))
                {
                    if ((txtdocDate.Text).Contains("/"))
                    {
                        foreach (DataGridViewRow gr in dgvPETC0303I.Rows)
                        {
                            gr.Cells["coldocDate"].Value = txtdocDate.Text;
                        }
                    }
                }
                else
                    txtdocDate.Focus();
            }
        }

        private void btnFolder_Click(object sender, EventArgs e)
        {
            DataTable dtPath = new DataTable();

            dtPath = PETC0309IBL.GetFolder3i();
            string folderPath = string.Empty;
            folderPath = dtPath.Rows[0]["Char2"].ToString();
            if (!string.IsNullOrWhiteSpace(folderPath))
            {
                if (!Directory.Exists(folderPath))
                {
                    Directory.CreateDirectory(folderPath);
                }
                OpenFileDialog op = new OpenFileDialog();
                op.Multiselect = true;
                op.InitialDirectory = folderPath;
                op.RestoreDirectory = true;
                op.Filter = "SMS files |*.csv;*.txt";
                if (op.ShowDialog() == DialogResult.OK)
                {
                    if (op.SafeFileNames.Count() == 1)
                    {
                        txtImport.Text = op.FileName;
                    }
                    else
                    {
                        txtImport.Text = op.SafeFileNames[0];
                    }
                    files = op.FileNames;
                }
            }
        }

        private void btnTorikomi_Click(object sender, EventArgs e)
        {
            F11();
        }

        private void dgvPETC0303I_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            dgcell = (sender as DataGridView).CurrentCell;  
        }

        private void dgvPETC0303I_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex != -1)   // Error skiiped when not clicking on header bar
            {
                if (dgvPETC0303I.CurrentCell.ColumnIndex != 0 && dgvPETC0303I.CurrentCell.ColumnIndex != 3 && dgvPETC0303I.CurrentCell.ColumnIndex != 1 && e.RowIndex >= 0)  // must not be these rows
                {
                    dgvPETC0303I.CurrentCell = dgvPETC0303I.Rows[e.RowIndex].Cells[e.ColumnIndex];
                    dgvPETC0303I.CurrentCell.Selected = true;  /// as soon as cursor point set the current cells, making it the cell edit
                    dgvPETC0303I.BeginEdit(true);  // make event for setting flags
                }
            }
        }

        private void dgvPETC0303I_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            Maintained_CellClick(sender, e);
        }

        private void dgvPETC0303I_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
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
                    if ((Convert.ToBoolean(dgvPETC0303I.Rows[e.RowIndex].Cells[e.ColumnIndex].EditedFormattedValue) == true))
                    {
                        MotoCD = dgvPETC0303I.Rows[e.RowIndex].Cells["colNyuukinMotoCD"].Value.ToString();
                        KouzaDate = dgvPETC0303I.Rows[e.RowIndex].Cells["colKouzaNyuukinDate"].Value.ToString();
                        string[] moto_kouza = { MotoCD, KouzaDate };
                        CheckALL((sender as DataGridView), "blur", moto_kouza, true);
                        foreach (DataGridViewRow row1 in PETC0303IDetail.Rows)
                        {
                            if (row1.Cells["colNyuukinMotoCD1"].Value.ToString() == MotoCD && row1.Cells["colKouzaNyuukinDate1"].Value.ToString() == KouzaDate)
                            {
                                res += Convert.ToInt32(row1.Cells["coltransferAmount"].Value.ToString().Replace(",", string.Empty));

                                row1.Cells["colchk"].Value = true;
                            }
                        }
                    }
                    else
                    {
                        MotoCD = dgvPETC0303I.Rows[e.RowIndex].Cells["colNyuukinMotoCD"].Value.ToString();
                        KouzaDate = dgvPETC0303I.Rows[e.RowIndex].Cells["colKouzaNyuukinDate"].Value.ToString();
                        string[] moto_kouza = { MotoCD, KouzaDate };
                        CheckALL((sender as DataGridView), "blur", moto_kouza, false);
                        foreach (DataGridViewRow row1 in PETC0303IDetail.Rows)
                        {
                            if (row1.Cells["colNyuukinMotoCD1"].Value.ToString() == MotoCD && row1.Cells["colKouzaNyuukinDate1"].Value.ToString() == KouzaDate)
                            {
                                row1.Cells["colchk"].Value = false;
                            }
                        }
                    }
                }

                foreach (DataGridViewRow row1 in dgvPETC0303I.Rows)
                {
                    if (row1.Cells["colNyuukinMotoCD"].Value.ToString() == MotoCD && row1.Cells["colKouzaNyuukinDate"].Value.ToString() == KouzaDate)
                    {
                        row1.Cells["col_nyuukingaku"].Value = res.ToString();
                        //   row1.Cells["col_nyuukingaku"].Value = CommaSeparate(res.ToString());
                    }
                }
            }
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

        private void dgvPETC0303I_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex >= (sender as DataGridView).Columns["colKeihi01input"].Index)
            {
                Sum_hiddenFiled(e.ColumnIndex);       //Enforeced ... make summation when ending in edit mode as a recover if someoneexceptions occurred
                dgvPETC0303I.NotifyCurrentCellDirty(true);   // put this not occuring win box errors for abnormal changing in cell values
            }
        }
        public void Sum_hiddenFiled(int colindex)  // make summation from all taxes calculated stored in hidded fields
        {
            decimal result = 0;
            var col_num = dgvPETC0303I.Columns[colindex].DataPropertyName.ToString().Replace("Keihi", "").Replace("Input", "").Replace("input", "");
            foreach (DataGridViewRow gr in dgvPETC0303I.Rows)
            {
                result = result + Convert.ToDecimal(gr.Cells["colKeihi" + col_num + "Input"].Value);
                if (gr.Index + 1 == dgvPETC0303I.Rows.Count)
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
        private bool CheckDate(string value)
        {
            DateTime d;
            var g = DateTime.TryParseExact(value, "yyyy/MM/dd", System.Globalization.CultureInfo.InvariantCulture, DateTimeStyles.None, out d);
            return g;
        }
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
        private void dgvPETC0303I_CellValidated(object sender, DataGridViewCellEventArgs e)
        {

            if ((sender as DataGridView).CurrentCell is DataGridViewTextBoxCell)
            {
                string chk_date = null;
                string[] dn = null;
                if ((Convert.ToBoolean(dgvPETC0303I.Rows[e.RowIndex].Cells["coltesu_chk"].EditedFormattedValue) == true) && (dgvPETC0303I.CurrentCell == dgvPETC0303I.Rows[e.RowIndex].Cells["coldocDate"]))
                {
                    if (dgvPETC0303I.Rows[e.RowIndex].Cells["coldocDate"].Value == null)
                    {
                        DSP_MSG("E102", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                    }
                    else
                    {
                        Date = dgvPETC0303I.Rows[e.RowIndex].Cells["coldocDate"].Value.ToString();
                        chk_date = checkDate(Date, dgvPETC0303I.CurrentCell);
                        dgvPETC0303I.Rows[e.RowIndex].Cells["coldocDate"].Value = chk_date;
                        if (chk_date != null)
                        {
                            MotoCD = dgvPETC0303I.Rows[e.RowIndex].Cells["colNyuukinMotoCD"].Value.ToString();
                            KouzaDate = dgvPETC0303I.Rows[e.RowIndex].Cells["colKouzaNyuukinDate"].Value.ToString();
                            string[] moto_kouza = { MotoCD, KouzaDate };
                            CheckALL((sender as DataGridView), "blur", moto_kouza, false, chk_date);
                        }
                    }
                }
                else if ((dgvPETC0303I.CurrentCell == dgvPETC0303I.Rows[e.RowIndex].Cells["coldocDate"]) && (!(dgvPETC0303I.Rows[e.RowIndex].Cells["coldocDate"].Value == null)))
                {
                    Date = dgvPETC0303I.Rows[e.RowIndex].Cells["coldocDate"].Value.ToString();
                    chk_date = checkDate(Date, dgvPETC0303I.CurrentCell);
                    dgvPETC0303I.Rows[e.RowIndex].Cells["coldocDate"].Value = chk_date;
                    if (chk_date != null)
                    {
                        MotoCD = dgvPETC0303I.Rows[e.RowIndex].Cells["colNyuukinMotoCD"].Value.ToString();
                        KouzaDate = dgvPETC0303I.Rows[e.RowIndex].Cells["colKouzaNyuukinDate"].Value.ToString();
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
                    // set_focus = true;
                }
                else  // set value tax and set as separated comma values
                {
                    string input = (sender as DataGridView).CurrentCell.EditedFormattedValue.ToString().Replace(",", "");
                    var tax_rate = dgvPETC0303I[(sender as DataGridView).CurrentCell.DataGridView.Columns["coldocDate"].Index, (sender as DataGridView).CurrentCell.RowIndex].Value;
                    var kbn_col = "colKeihi" + (sender as DataGridView).Columns[(sender as DataGridView).CurrentCell.ColumnIndex].DataPropertyName.ToString().Replace("Keihi", "").Replace("Input", "").Replace("input", "") + "TaxKBN";
                    var kbn = dgvPETC0303I.Rows[(sender as DataGridView).CurrentCell.RowIndex].Cells[kbn_col].Value;
                    var tax_col = "colKeihi" + (sender as DataGridView).Columns[(sender as DataGridView).CurrentCell.ColumnIndex].DataPropertyName.ToString().Replace("Keihi", "").Replace("Input", "") + "Tax";
                    (sender as DataGridView)[tax_col, (sender as DataGridView).CurrentCell.RowIndex].Value = Calculate_Tax(kbn.ToString(), GetTaxtRate(dgvPETC0303I.Rows[e.RowIndex].Cells["coldocDate"].EditedFormattedValue), input.ToString());
                    (sender as DataGridView).CurrentCell.Value = CommaSeparate(input);
                }
            }
        }
        public static object CommaSeparate(string value)  // return comma separated values
        {
            object j = Round_Off(value).ToString("#,##0");
            return j;
        }

        private void dgvPETC0303I_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            try
            {
                (sender as DataGridView).EditingControl.ResetText();
                (sender as DataGridView).EditingControl.Text = "0";
                dgvPETC0303I.CellEndEdit += dgvPETC0303I_CellEndEdit;
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
            catch { }
        }
        public void BindGrid_Sorting()
        {


            for (int c = 1; c <= 15; c++)
            {
                decimal result = 0;
                foreach (DataGridViewRow gr in dgvPETC0303I.Rows)
                {
                    var val = (c).ToString().PadLeft(2, '0');
                    //   gr.Cells["colKeihi" + val + "Tax"].Value = Calculate_Tax(gr.Cells["colKeihi" + val + "TaxKBN"].Value.ToString(), gr.Cells["colTaxRate"].Value.ToString(), gr.Cells["colKeihi" + val + "input"].Value.ToString().Replace(",", ""));
                    gr.Cells["colKeihi" + val + "Tax"].Value = Calculate_Tax(gr.Cells["colKeihi" + val + "TaxKBN"].Value.ToString(), GetTaxtRate(gr.Cells["coldocDate"].EditedFormattedValue), gr.Cells["colKeihi" + val + "input"].Value.ToString().Replace(",", ""));
                    result = result + Convert.ToDecimal(gr.Cells["colKeihi" + val + "Input"].Value);
                    if (gr.Index + 1 == dgvPETC0303I.Rows.Count)
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

            //foreach (DataGridViewRow gr in dgvPETC0303I.Rows)
            //{
            //    gr.Cells["col_nyuukingaku"].Value = Convert.ToDecimal(gr.Cells["col_nyuukingaku"].Value).ToString("#,##0");
            //}




        }
        private void dgvPETC0303I_Sorted(object sender, EventArgs e)
        {
            try
            {
                BindGrid_Sorting();
            }
            catch { }
        }

        private void PETC0303IDetail_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            MaintainCellClick_detail(sender, e);
        }

        private void PETC0303IDetail_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            MaintainCellClick_detail(sender, e);
        }
        private void MaintainCellClick_detail(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
                if ((sender as DataGridView).CurrentCell is DataGridViewCheckBoxCell)
                {

                    if ((Convert.ToBoolean(PETC0303IDetail.Rows[e.RowIndex].Cells[e.ColumnIndex].EditedFormattedValue) == true))
                    {
                        MotoCD = PETC0303IDetail.Rows[e.RowIndex].Cells["colNyuukinMotoCD1"].Value.ToString();
                        KouzaDate = PETC0303IDetail.Rows[e.RowIndex].Cells["colKouzaNyuukinDate1"].Value.ToString();
                        string[] moto_kouza = { MotoCD, KouzaDate };
                        CheckALL(dgvPETC0303I, "blur", moto_kouza, true);

                    }
                    else
                    {

                        // this.dgvPETC0302I.Refresh();
                        bool flg = false;
                        foreach (DataGridViewRow row1 in PETC0303IDetail.Rows)
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
                            MotoCD = PETC0303IDetail.Rows[e.RowIndex].Cells["colNyuukinMotoCD1"].Value.ToString();
                            KouzaDate = PETC0303IDetail.Rows[e.RowIndex].Cells["colKouzaNyuukinDate1"].Value.ToString();
                            string[] moto_kouza = { MotoCD, KouzaDate };
                            CheckALL(dgvPETC0303I, "blur", moto_kouza, false);
                        }

                    }


                    MotoCD = PETC0303IDetail.Rows[e.RowIndex].Cells["colNyuukinMotoCD1"].Value.ToString();
                    KouzaDate = PETC0303IDetail.Rows[e.RowIndex].Cells["colKouzaNyuukinDate1"].Value.ToString();
                    int res = 0;
                    foreach (DataGridViewRow row1 in PETC0303IDetail.Rows)
                    {
                        if (row1.Cells["colNyuukinMotoCD1"].Value.ToString() == MotoCD && row1.Cells["colKouzaNyuukinDate1"].Value.ToString() == KouzaDate)
                        {
                            if ((Convert.ToBoolean(row1.Cells["colchk"].EditedFormattedValue) == true))
                            {
                                res += Convert.ToInt32(row1.Cells["coltransferAmount"].EditedFormattedValue.ToString().Replace(",", ""));
                            }
                        }
                    }
                    foreach (DataGridViewRow row1 in dgvPETC0303I.Rows)
                    {
                        if (row1.Cells["colNyuukinMotoCD"].Value.ToString() == MotoCD && row1.Cells["colKouzaNyuukinDate"].Value.ToString() == KouzaDate)
                        {
                            row1.Cells["col_nyuukingaku"].Value = res.ToString();
                            //   row1.Cells["col_nyuukingaku"].Value = CommaSeparate(res.ToString());
                            //  res += Convert.ToInt32(row1.Cells["colNyuukinGaku"].EditedFormattedValue.ToString().Replace(",", ""));

                        }
                    }
                }
        }

        private void PETC0303IDetail_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex >= 0)
                if ((sender as DataGridView).CurrentCell is DataGridViewCheckBoxCell)
                {

                    if ((Convert.ToBoolean(PETC0303IDetail.Rows[e.RowIndex].Cells[e.ColumnIndex].EditedFormattedValue) == true))
                    {
                        MotoCD = PETC0303IDetail.Rows[e.RowIndex].Cells["colNyuukinMotoCD1"].Value.ToString();
                        KouzaDate = PETC0303IDetail.Rows[e.RowIndex].Cells["colKouzaNyuukinDate1"].Value.ToString();
                        string[] moto_kouza = { MotoCD, KouzaDate };
                        CheckALL(dgvPETC0303I, "blur", moto_kouza, true);

                    }
                    else
                    {
                        // this.dgvPETC0302I.Refresh();
                        bool flg = false;
                        foreach (DataGridViewRow row1 in PETC0303IDetail.Rows)
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
                            MotoCD = PETC0303IDetail.Rows[e.RowIndex].Cells["colNyuukinMotoCD1"].Value.ToString();
                            KouzaDate = PETC0303IDetail.Rows[e.RowIndex].Cells["colKouzaNyuukinDate1"].Value.ToString();
                            string[] moto_kouza = { MotoCD, KouzaDate };
                            CheckALL(dgvPETC0303I, "blur", moto_kouza, false);
                        }

                    }
                }
        }

        private void btnAllCheckon_Click(object sender, EventArgs e)
        {
            if (!String.IsNullOrEmpty(txtImport.Text))
            {
                BindGrid(txtdocDate.Text);
            }
            CheckALL(dgvPETC0303I, "coltesu_chk", null, true, null);
            CheckALL(PETC0303IDetail, "colchk", null, true, null);
        }

        private void btnAllCheckOff_Click(object sender, EventArgs e)
        {

            if (!String.IsNullOrEmpty(txtImport.Text))
            {
                BindGrid(txtdocDate.Text);
            }

            CheckALL(dgvPETC0303I, "coltesu_chk", null, false, null);
            CheckALL(PETC0303IDetail, "colchk", null, false, null);

            foreach (DataGridViewRow gr in dgvPETC0303I.Rows)
            {
                gr.Cells["col_nyuukingaku"].Value = "0";
            }
        }
    }
}
