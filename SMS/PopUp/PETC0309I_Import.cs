using System;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Collections;
using System.Threading;
using System.IO;
using ExcelDataReader;
using Microsoft.VisualBasic.FileIO;
using SMS_BL;
using SMS_Entity;

namespace SMS.PopUp
{
    public partial class PETC0309I_Import : Form
    {
        string firstColHeader = string.Empty;
        Base_BL baseBL;
        T_NyuukinMeisai_Entity NyuukinMeisai_data;
        CommonFunctions cmf;
        PETC0309I_BL PETC30303I;
        L_Log_Entity LLogEntity;
        string[] arraylist;
        public PETC0309I_Import(string[] arl, L_Log_Entity lle)
        {
            InitializeComponent(); 
            this.LLogEntity = lle;
            this.arraylist = arl;
            BindGrid(arraylist);
            dgv1.DisabledColumn("colNo");
        }
        private DataTable CSVToDataTable(string filePathName)
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
        private async void PETC0309I_Import_Load(object sender, EventArgs e)
        {
            baseBL = new Base_BL();
            NyuukinMeisai_data = new T_NyuukinMeisai_Entity();
            PETC30303I = new PETC0309I_BL();
            cmf = new CommonFunctions();
            //  psks0103ibl = new PSKS0103I_BL();
            btnOK.Enabled = false;

            if (PETC30303I.PETC0309IDelete(LLogEntity))
            {
                for (int i = 0; i < dgv1.Rows.Count; i++)
                {
                    await Insert(i);
                }
            }

            btnOK.Enabled = true;
        }
        public async Task<bool> Insert(int i)
        {
            Func<bool> function = new Func<bool>(() => InsertData(i));
            return await Task.Factory.StartNew<bool>(function);
        }
        public string GetFirst2Kanji(string filesname)
        {
            byte[] unicodeBytes = Encoding.UTF32.GetBytes(filesname);
            try
            {
                return Encoding.UTF32.GetString(unicodeBytes, 0, 2 * 4); // 4 = number of bytes per character
            }
            catch
            {
                return null;
            }

        }

        private void BindGrid(string[] arrlst)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("No");
            dt.Columns.Add("FileName");
            dt.Columns.Add("Status");

            for (int i = 0; i < arrlst.Length; i++)
            {
                DataRow dr = dt.NewRow();
                dr["No"] = (i + 1).ToString();
                dr["FileName"] = arrlst[i];
                dr["Status"] = "Waiting.....";
                dt.Rows.Add(dr);
            }

            dgv1.DataSource = dt;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        public DataTable TXTToDataTable(string filePath, int numberOfColumns = 11)    // for error Check E137 
        {
            DataTable tbl = new DataTable();
            string[] header = { "入金日付", "コード", "得意先", "振込", "現金", "小切手", "手形", "相殺", "調整", "入金(合計)", "備考" };
            //for (int o = 0; o < tbl.Rows.Count; o++)
            //{
            //    tbl.Columns[""].ColumnName = "";

            //}
            for (int col = 0; col < numberOfColumns; col++)   // ignore First Line header
                tbl.Columns.Add(new DataColumn(header[col]).ToString());

            string[] lines = System.IO.File.ReadAllLines(filePath, Encoding.GetEncoding(932));
            //Check First is Not 
            int cout = 0;
            foreach (string line in lines)
            {
                cout++;
                if (cout == 1)
                {
                    var cols = line.Split(',');

                    if (cols[0].Replace("\"", string.Empty) != "入金日計明細表")  // Return Error if not this text
                    {
                        cmf.DSP_MSG("E137", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                        return null;
                    }

                }
                else if (cout != 1 && cout != 2 && cout != 3)
                {
                    var cols = (new System.Text.RegularExpressions.Regex("\\\"(.*?)\\\"")).Replace(line, m => m.Value.Replace(",", string.Empty)).Split(',');



                    if (cols[0].Replace("\"", string.Empty) != "合計")     // Check  Not date
                    {
                        DataRow dr = tbl.NewRow();
                        for (int cIndex = 0; cIndex < cols.Count(); cIndex++)
                        {
                            if (cIndex == 0)
                            {
                                if (cols.Count() != 1)
                                dr[cIndex] = Convert.ToDateTime(cols[cIndex].Replace("\"", "")).ToString();
                            }
                            else
                            {
                                dr[cIndex] = cols[cIndex].Replace("\"", "");
                            }
                        }

                        tbl.Rows.Add(dr);
                    }

                }
            }

            return tbl;
        }
        public bool InsertData(int i)
        {

            dgv1.Rows[i].Cells["colStatus"].Value = "Importing.....";
            dgv1.Rows[i].Cells["colImage"].Value = SMS.Properties.Resources.dbImport;
            // dgv1.Rows[i].Cells["colImage"].Style.BackColor = Color.FromArgb(255, 255, 192);
            dgv1.Rows[i].Cells["colReason"].Value = ".....";
            dgv1.Rows[i].DefaultCellStyle.BackColor = Color.WhiteSmoke;
            dgv1.Rows[i].DefaultCellStyle.ForeColor = Color.Black;
            string filePath = dgv1.Rows[i].Cells["colFileName"].Value.ToString();
            string ext = Path.GetExtension(filePath);

            var dtImport = new DataTable();
            if (ext.Equals(".csv") || ext.Equals(".CSV"))
            {
                dtImport = CSVToDataTable(filePath);  ///Wrong csv to datatable 
            }
            else if (ext.Equals(".txt"))
            {
                dtImport = TXTToDataTable(filePath);
            }
            else
            {
                dgv1.Rows[i].Cells["colStatus"].Value = "Invalid File";
                dgv1.Rows[i].Cells["colImage"].Value = SMS.Properties.Resources.close;
                dgv1.Rows[i].Cells["colReason"].Value = "Choose correct format File!";
                dgv1.Rows[i].DefaultCellStyle.BackColor = Color.Red;
                dgv1.Rows[i].DefaultCellStyle.ForeColor = Color.White;
                return false;
            }
            if (dtImport != null)
            {
                NyuukinMeisai_data.xml = baseBL.DataTableToXml(dtImport);
                var check2kanji = GetFirst2Kanji(Path.GetFileNameWithoutExtension(filePath));
                if (check2kanji != null)   // Chech First 2 kanji
                {
                    if (PETC30303I.CheckExistName1_0309I(NyuukinMeisai_data, LLogEntity, check2kanji))
                    {
                        if (PETC30303I.PETC0309I_Insert(NyuukinMeisai_data, LLogEntity, check2kanji))    // Meisai and tetsuyou
                        {
                            dgv1.Rows[i].Cells["colStatus"].Value = "Finished!";
                            dgv1.Rows[i].Cells["colImage"].Value = SMS.Properties.Resources.success;
                            dgv1.Rows[i].Cells["colReason"].Value = dtImport.Rows.Count.ToString() + " data have been imported!";
                            dgv1.Rows[i].DefaultCellStyle.BackColor = Color.FromArgb(192, 255, 192);
                            dgv1.Rows[i].DefaultCellStyle.ForeColor = Color.Black;
                        }
                    }
                    else
                    {
                        dgv1.Rows[i].Cells["colStatus"].Value = "Failed!";
                        dgv1.Rows[i].Cells["colImage"].Value = SMS.Properties.Resources.close;
                        cmf.DSP_MSG("E137", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty); // Error While not Catching first 2 kanjis
                        dgv1.Rows[i].Cells["colReason"].Value = "No row exist in import field with respect to NyuukinmotoName1!";
                        dgv1.Rows[i].DefaultCellStyle.BackColor = Color.CornflowerBlue;
                        dgv1.Rows[i].DefaultCellStyle.ForeColor = Color.White;
                    }
                }
                else
                {

                    dgv1.Rows[i].Cells["colStatus"].Value = "Failed!";
                    dgv1.Rows[i].Cells["colImage"].Value = SMS.Properties.Resources.close;
                    cmf.DSP_MSG("E137", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty); // Error While not Catching first 2 kanjis
                    dgv1.Rows[i].Cells["colReason"].Value = "Error in getting First two Kanji!";
                    dgv1.Rows[i].DefaultCellStyle.BackColor = Color.CornflowerBlue;
                    dgv1.Rows[i].DefaultCellStyle.ForeColor = Color.White;
                }
            }
            else
            {
                dgv1.Rows[i].Cells["colStatus"].Value = "Failed!";
                dgv1.Rows[i].Cells["colImage"].Value = SMS.Properties.Resources.close;
                dgv1.Rows[i].Cells["colReason"].Value = "Not importable data file(No data)!";
                dgv1.Rows[i].DefaultCellStyle.BackColor = Color.CornflowerBlue;
                dgv1.Rows[i].DefaultCellStyle.ForeColor = Color.White;
            }
            dgv1.Rows[i].Cells["colNo"].Style.ForeColor = Color.Black;
            return true;
        }

        private void btnOK_Click_1(object sender, System.EventArgs e)
        {
            this.Close();
        }
    }
}
