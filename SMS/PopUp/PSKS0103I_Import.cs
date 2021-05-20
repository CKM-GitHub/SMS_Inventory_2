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
using System.Collections;
using System.Threading;
using System.IO;
using ExcelDataReader;
using Microsoft.VisualBasic.FileIO;
using SMS_BL;

namespace SMS.PopUp
{
    public partial class PSKS0103I_Import : Form
    {

        ArrayList arrlst;
        T_MakerZaiko_Entity tmze;
        PSKS0103I_BL psks0103ibl;
        string firstColHeader = string.Empty;
        string patternCD = string.Empty;
        MakerZaiko_D_Entity mzde;
        public PSKS0103I_Import(ArrayList arrlst, T_MakerZaiko_Entity tmze, string firstColHeader, string patternCD)
        {
            InitializeComponent();
            this.arrlst = arrlst;
            this.tmze = tmze;
            lblNum.Text = arrlst.Count.ToString();
            this.firstColHeader = firstColHeader;
            this.patternCD = patternCD;
            BindGrid(arrlst);
            //dgv1.AlternatingRowsDefaultCellStyle.BackColor = Color.Empty;
            dgv1.DisabledColumn("colNo");
        }

        private void BindGrid(ArrayList arrlst)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("No");
            dt.Columns.Add("FileName");
            dt.Columns.Add("Status");

            for (int i = 0; i < arrlst.Count; i++)
            {
                DataRow dr = dt.NewRow();
                dr["No"] = (i + 1).ToString();
                dr["FileName"] = arrlst[i];
                dr["Status"] = "Waiting.....";

                dt.Rows.Add(dr);
            }

            dgv1.DataSource = dt;
        }

        private async void PSKS0103I_Import_Load(object sender, EventArgs e)
        {
            psks0103ibl = new PSKS0103I_BL();
            btnOK.Enabled = false;

            for (int i = 0; i < dgv1.Rows.Count; i++)
            {
                await Insert(i);
            }

            btnOK.Enabled = true;
        }

        public async Task<bool> Insert(int i)
        {
            Func<bool> function = new Func<bool>(() => InsertData(i));
            return await Task.Factory.StartNew<bool>(function);
        }

        private bool CheckAllPhyoeGyiColumns(DataTable dt, string Keys)
        {
            int h = 0;
            foreach (DataColumn dc in dt.Columns)
            {
                if (dc.ColumnName.Contains(Keys))
                {
                    h++;
                }
            }

            return h.ToString() == dt.Columns.Count.ToString();
        }
        public bool InsertData(int i)
        {
            try
            {
                dgv1.Rows[i].Cells["colStatus"].Value = "Importing.....";
                dgv1.Rows[i].Cells["colImage"].Value = SMS.Properties.Resources.dbImport;
                dgv1.Rows[i].Cells["colImage"].Style.BackColor = Color.FromArgb(255, 255, 192);


                string filePath = dgv1.Rows[i].Cells["colFileName"].Value.ToString();
                string ext = Path.GetExtension(filePath);

                DataTable dtImport = new DataTable();

                mzde = new MakerZaiko_D_Entity();
                mzde.PatternCD = patternCD;
                DataTable dtPattern = psks0103ibl.M_MakerZaiko_D_Select(mzde);

                if (ext.Equals(".csv"))
                {
                    if (patternCD == "011") /// for sankyo_ maker by etz
                        dtImport = CSVToTable(filePath, dtPattern.Rows.Count);
                    else
                        dtImport = CSVToTable(filePath);
                }
                else if (ext.Equals(".xlsx") || ext.Equals(".xls"))
                {
                    dtImport = ExcelToDatatable(filePath);
                }
                else
                {
                    dgv1.Rows[i].Cells["colStatus"].Value = "Invalid File";
                    dgv1.Rows[i].Cells["colImage"].Value = SMS.Properties.Resources.close;
                    dgv1.Rows[i].Cells["colImage"].Style.BackColor = Color.Red;
                    return false;
                }

                //Delete Extra Column With Blank Rows   ---ktp 2019-04-16
                ArrayList arrdatacolumn = new ArrayList();

                string Keys = "MaungPhyoeGyi_JennieKim";  // -------PTK Added Use-header row Conflict
                if (!CheckAllPhyoeGyiColumns(dtImport, Keys)) /// PTK Added For Use-header row Conflict
                {
                    foreach (DataColumn dc in dtImport.Columns)
                    {
                        if (dc.ColumnName.Contains(Keys))
                        {
                            arrdatacolumn.Add(dc);
                        }
                    }

                    foreach (DataColumn dc in arrdatacolumn)
                    {
                        dtImport.Columns.Remove(dc);
                    }
                }
                dtImport.AcceptChanges();
                try
                {
                    M_MultiPorpose_Entity mme1 = new M_MultiPorpose_Entity();
                    mme1.fields = "Char1,Num1,Char5";
                    mme1.tableName = "M_MultiPorpose";
                    mme1.condition = "Where ID=101 And [Key]='1'";

                    DataTable dt1 = psks0103ibl.M_MultiPorpose_DynamicSelect(mme1);
                    if (dt1.Rows[0]["Char5"].ToString() == "1")
                    {
                        MessageBox.Show("Excel cols " + dtImport.Columns.Count.ToString() + Environment.NewLine + "Pattern Cols" + dtPattern.Rows.Count.ToString());
                    }
                }
                catch { }
                if (dtImport.Columns.Count != dtPattern.Rows.Count)
                {
                    DSP_MSG("E137", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                    dgv1.Rows[i].Cells["colStatus"].Value = "Failed";
                    dgv1.Rows[i].Cells["colImage"].Value = SMS.Properties.Resources.Failed;
                    dgv1.Rows[i].Cells["colImage"].Style.BackColor = Color.Red;

                    return false;
                }

                tmze.dt1 = dtImport;
                tmze.ImportFileName = Path.GetFileName(filePath);

                M_MultiPorpose_Entity mme = new M_MultiPorpose_Entity();
                mme.fields = "Char1,Char2,Num1";
                mme.tableName = "M_MultiPorpose";
                mme.condition = "Where ID=101 And [Key]='1'";

                DataTable dt = psks0103ibl.M_MultiPorpose_DynamicSelect(mme);

                //string destination = @"C:\SMS\PSKS0103I\ImportFinished\";

                string destination = String.IsNullOrEmpty(dt.Rows[0]["Char2"].ToString()) ? @"C:\SMS\PSKS0103I\ImportFinished\" : dt.Rows[0]["Char2"].ToString() + @"\";

                if (psks0103ibl.ImportProcess(tmze, patternCD, dtPattern, i, System.DateTime.Now.ToString()))
                {

                    dgv1.Rows[i].Cells["colStatus"].Value = "Finished";
                    dgv1.Rows[i].Cells["colImage"].Value = SMS.Properties.Resources.success;
                    dgv1.Rows[i].Cells["colImage"].Style.BackColor = Color.FromArgb(192, 255, 192);
                    if (!Directory.Exists(destination))
                    {
                        Directory.CreateDirectory(destination);
                    }

                    if (File.Exists(destination + Path.GetFileName(filePath)))
                        File.Delete(destination + Path.GetFileName(filePath));
                    try
                    {
                        File.Move(filePath, destination + Path.GetFileName(filePath));
                    }
                    catch { };
                }
                else
                {

                    dgv1.Rows[i].Cells["colStatus"].Value = "Failed";
                    dgv1.Rows[i].Cells["colImage"].Value = SMS.Properties.Resources.close;
                    dgv1.Rows[i].Cells["colImage"].Style.BackColor = Color.Red;
                }

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.InnerException.ToString());
                return false;
            }
        }
        public string DSP_MSG(string messageID, string messageParam1, string messageParam2, string messageParam3, string messageParam4, string messageParam5)
        {
            CommonFunctions cf = new CommonFunctions();
            return cf.DSP_MSG(messageID, messageParam1, messageParam2, messageParam3, messageParam4, messageParam5);
        }
        public bool IsColumnCountOKay(string Path, string mkr_CD, int count)
        {

            psks0103ibl = new PSKS0103I_BL();

            DataTable dtpthname = psks0103ibl.get_ptnname(mkr_CD);

            //var e_2_d = ExcelToDatatable(Path);

            if (Convert.ToInt32(dtpthname.Rows[0]["Total_column"]).Equals(count))
            {
                return true;
            }
            else
            {
                return false;
            }
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
            if (firstColHeader.Equals("2"))
                useHeaderRow = false;

            DataSet result = excelReader.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = useHeaderRow,
                    EmptyColumnNamePrefix= "MaungPhyoeGyi_JennieKim",

                }
            });


            excelReader.Close();
            return result.Tables[0];
        }

        public DataTable CSVToTable(string filePath, int COL_COUNT = 0)
        {
            DataTable csvData = new DataTable();
            int count = 1;
            try
            {
                //Array lines = File.ReadAllLines(filePath);
                //var file = File
                //    .ReadAllLines(filePath, Encoding.GetEncoding(932))
                //    .SkipWhile(line => string.IsNullOrWhiteSpace(line)) // To be on the safe side
                //    .Select((line, index) => index == 1 // do we have header? 
                //        ? line.Replace('\"',' ') // replace '_' with ' '
                //        : line)                  // keep lines as they are 
                //    .ToList();                  // Materialization, since we write into the same file

                //File.WriteAllLines(filePath, file, Encoding.GetEncoding(932));

                using (TextFieldParser csvReader = new TextFieldParser(filePath, Encoding.GetEncoding(932), true))
                {
                    csvReader.SetDelimiters(new string[] { "," });
                    csvReader.HasFieldsEnclosedInQuotes = true;
                    //read column names
                    string[] colFields = csvReader.ReadFields();
                    char c = 'A';

                    //for (int i = 0; i < colFields.Length; i++)
                    //{
                    //    colFields[i] = colFields[i].Replace("\"", "");
                    //}

                    if (COL_COUNT != 0 && colFields.Length != COL_COUNT)
                    {
                        int ct = colFields.Count();
                        while (colFields.Count() > ct - 4)
                        {
                            colFields = colFields.Take(colFields.Count() - 1).ToArray();
                        }


                    }

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
                            //if(!csvData.Columns.Contains("\"\"") || !csvData.Columns.Contains(""))
                            //{
                            DataColumn datacolumn = new DataColumn(c++.ToString());
                            datacolumn.AllowDBNull = true;
                            csvData.Columns.Add(datacolumn);
                            // }

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
                        csvData.Rows.Add(fieldData);

                    }
                }
            }
            catch (Exception ex)
            {
            }
            return csvData;
        }

        public static int CountStringOccurrences(string[] arr, String item)
        {
            int i; int j = 0;
            for (i = 0; i < arr.Count(); i++)
            {
                if (!string.IsNullOrEmpty(arr[i]))
                {
                    if (arr[i].Contains(item))
                    {
                        j++;
                    }
                }
            }
            return j;
        }

        public static DataTable CsvToDataTable(string FilePath)
        {
            DataTable csvData = new DataTable();
            try
            {
                using (TextFieldParser csvReader = new TextFieldParser(FilePath, Encoding.GetEncoding(932)))
                {
                    csvReader.SetDelimiters(new string[] { "," });

                    csvReader.HasFieldsEnclosedInQuotes = true;

                    string[] colFields = csvReader.ReadFields();

                    foreach (string column in colFields)
                    {

                        DataColumn datecolumn = new DataColumn(column);

                        datecolumn.AllowDBNull = true;

                        csvData.Columns.Add(datecolumn);
                    }

                    while (!csvReader.EndOfData)
                    {
                        string[] fieldData = csvReader.ReadFields();

                        //Making empty value as null

                        for (int i = 0; i < fieldData.Length; i++)
                        {
                            if (fieldData[i] == "")
                            {
                                fieldData[i] = null;
                            }
                        }
                        csvData.Rows.Add(fieldData);
                    }
                }
            }

            catch (Exception)
            {

            }
            return csvData;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
//using System;
//using System.Collections.Generic;
//using System.ComponentModel;
//using System.Data;
//using System.Drawing;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
//using System.Windows.Forms;
//using SMS_Entity;
//using System.Collections;
//using System.Threading;
//using System.IO;
//using ExcelDataReader;
//using Microsoft.VisualBasic.FileIO;
//using SMS_BL;

//namespace SMS.PopUp
//{
//    public partial class PSKS0103I_Import : Form
//    {
//        ArrayList arrlst;
//        T_MakerZaiko_Entity tmze;
//        PSKS0103I_BL psks0103ibl;
//        string firstColHeader = string.Empty;
//        string patternCD = string.Empty;
//        MakerZaiko_D_Entity mzde;
//        public PSKS0103I_Import(ArrayList arrlst, T_MakerZaiko_Entity tmze, string firstColHeader,string patternCD)
//        {
//            InitializeComponent();
//            this.arrlst = arrlst; 
//            this.tmze = tmze;
//            lblNum.Text = arrlst.Count.ToString();
//            this.firstColHeader = firstColHeader;
//            this.patternCD = patternCD;
//            BindGrid(arrlst);
//            //dgv1.AlternatingRowsDefaultCellStyle.BackColor = Color.Empty;
//            dgv1.DisabledColumn("colNo");
//        }

//        private void BindGrid(ArrayList arrlst)
//        {
//            DataTable dt = new DataTable();
//            dt.Columns.Add("No");
//            dt.Columns.Add("FileName");
//            dt.Columns.Add("Status");

//            for (int i = 0; i < arrlst.Count; i++)
//            {
//                DataRow dr = dt.NewRow();
//                dr["No"] = (i + 1).ToString();
//                dr["FileName"] = arrlst[i];
//                dr["Status"] = "Waiting.....";

//                dt.Rows.Add(dr);
//            }

//            dgv1.DataSource = dt;
//        }

//        private async void PSKS0103I_Import_Load(object sender, EventArgs e)
//        {
//            psks0103ibl = new PSKS0103I_BL();
//            btnOK.Enabled = false;

//            for (int i = 0; i < dgv1.Rows.Count; i++)
//            {
//                await Insert(i);
//            }

//            btnOK.Enabled = true;
//        }

//        public async Task<bool> Insert(int i)
//        {
//            Func<bool> function = new Func<bool>(() => InsertData(i));
//            return await Task.Factory.StartNew<bool>(function);
//        }


//        public bool InsertData(int i)
//        {
//            try
//            {
//                dgv1.Rows[i].Cells["colStatus"].Value = "Importing.....";
//                dgv1.Rows[i].Cells["colImage"].Value = SMS.Properties.Resources.dbImport;
//                dgv1.Rows[i].Cells["colImage"].Style.BackColor = Color.FromArgb(255, 255, 192);


//                string filePath = dgv1.Rows[i].Cells["colFileName"].Value.ToString();
//                string ext = Path.GetExtension(filePath);

//                DataTable dtImport = new DataTable();

//                mzde = new MakerZaiko_D_Entity();
//                mzde.PatternCD = patternCD;
//                DataTable dtPattern = psks0103ibl.M_MakerZaiko_D_Select(mzde);

//                if (ext.Equals(".csv"))
//                {
//                    if (patternCD=="011") /// for sankyo_ maker by etz
//                     dtImport = CSVToTable(filePath, dtPattern.Rows.Count);
//                    else
//                        dtImport = CSVToTable(filePath);
//                }
//                else if (ext.Equals(".xlsx") || ext.Equals(".xls"))
//                {
//                    dtImport = ExcelToDatatable(filePath);
//                }
//                else
//                {
//                    dgv1.Rows[i].Cells["colStatus"].Value = "Invalid File";
//                    dgv1.Rows[i].Cells["colImage"].Value = SMS.Properties.Resources.close;
//                    dgv1.Rows[i].Cells["colImage"].Style.BackColor = Color.Red;
//                    return false;
//                }

//                //Delete Extra Column With Blank Rows   ---ktp 2019-04-16
//                ArrayList arrdatacolumn = new ArrayList();

//                foreach (DataColumn dc in dtImport.Columns)
//                {
//                    if (dc.ColumnName.Contains("Column"))
//                    {
//                        arrdatacolumn.Add(dc);
//                    }
//                }

//                foreach (DataColumn dc in arrdatacolumn)
//                {
//                    dtImport.Columns.Remove(dc);
//                }
//                dtImport.AcceptChanges();



//                if (dtImport.Columns.Count != dtPattern.Rows.Count)
//                {
//                    DSP_MSG("E137", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
//                    dgv1.Rows[i].Cells["colStatus"].Value = "Failed";
//                    dgv1.Rows[i].Cells["colImage"].Value = SMS.Properties.Resources.Failed;
//                    dgv1.Rows[i].Cells["colImage"].Style.BackColor = Color.Red;

//                    return false;
//                }

//                tmze.dt1 = dtImport;
//                tmze.ImportFileName = Path.GetFileName(filePath);

//                M_MultiPorpose_Entity mme = new M_MultiPorpose_Entity();
//                mme.fields = "Char1,Char2,Num1";
//                mme.tableName = "M_MultiPorpose";
//                mme.condition = "Where ID=101 And [Key]='1'";

//                DataTable dt = psks0103ibl.M_MultiPorpose_DynamicSelect(mme);

//                //string destination = @"C:\SMS\PSKS0103I\ImportFinished\";

//                string destination = String.IsNullOrEmpty(dt.Rows[0]["Char2"].ToString()) ? @"C:\SMS\PSKS0103I\ImportFinished\" : dt.Rows[0]["Char2"].ToString() + @"\";

//                if (psks0103ibl.ImportProcess(tmze,patternCD,dtPattern, i, System.DateTime.Now.ToString()))
//                {

//                    dgv1.Rows[i].Cells["colStatus"].Value = "Finished";
//                    dgv1.Rows[i].Cells["colImage"].Value = SMS.Properties.Resources.success;
//                    dgv1.Rows[i].Cells["colImage"].Style.BackColor = Color.FromArgb(192, 255, 192);
//                    if (!Directory.Exists(destination))
//                    {
//                        Directory.CreateDirectory(destination);
//                    }

//                    if (File.Exists(destination + Path.GetFileName(filePath)))
//                        File.Delete(destination + Path.GetFileName(filePath));
//                    File.Move(filePath, destination + Path.GetFileName(filePath));
//                }
//                else
//                {

//                    dgv1.Rows[i].Cells["colStatus"].Value = "Failed";
//                    dgv1.Rows[i].Cells["colImage"].Value = SMS.Properties.Resources.close;
//                    dgv1.Rows[i].Cells["colImage"].Style.BackColor = Color.Red;
//                }

//                return true;
//            }
//            catch (Exception ex)
//            {
//                MessageBox.Show(ex.InnerException.ToString());
//                return false;
//            }
//        }
//        public string DSP_MSG(string messageID, string messageParam1, string messageParam2, string messageParam3, string messageParam4, string messageParam5)
//        {
//            CommonFunctions cf = new CommonFunctions();
//            return cf.DSP_MSG(messageID, messageParam1, messageParam2, messageParam3, messageParam4, messageParam5);
//        }
//        public bool IsColumnCountOKay(string Path, string mkr_CD, int count)
//        {

//            psks0103ibl = new PSKS0103I_BL();

//            DataTable dtpthname = psks0103ibl.get_ptnname(mkr_CD);

//            //var e_2_d = ExcelToDatatable(Path);

//            if (Convert.ToInt32(dtpthname.Rows[0]["Total_column"]).Equals(count))
//            {
//                return true;
//            }
//            else
//            {
//                return false;
//            }
//        }

//        protected DataTable ExcelToDatatable(string filePath)
//        {
//            FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read);

//            string ext = Path.GetExtension(filePath);
//            IExcelDataReader excelReader;
//            if (ext.Equals(".xls"))
//                //1. Reading from a binary Excel file ('97-2003 format; *.xls)
//                excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
//            else if (ext.Equals(".xlsx"))
//                //2. Reading from a OpenXml Excel file (2007 format; *.xlsx)
//                excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
//            else
//                //2. Reading from a OpenXml Excel file (2007 format; *.xlsx) 
//                excelReader = ExcelReaderFactory.CreateCsvReader(stream, null);

//            //3. DataSet - The result of each spreadsheet will be created in the result.Tables
//            bool useHeaderRow = true;
//            if (firstColHeader.Equals("2"))
//                useHeaderRow = false;

//            DataSet result = excelReader.AsDataSet(new ExcelDataSetConfiguration()
//            {
//                ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
//                {                   
//                    UseHeaderRow = useHeaderRow,
//                }
//            });


//            excelReader.Close();
//            return result.Tables[0];
//        }

//        public DataTable CSVToTable(string filePath,int COL_COUNT=0)
//        {
//            DataTable csvData = new DataTable();
//            int count = 1;
//            try
//            {
//                int row = System.IO.File.ReadAllLines(filePath).Length;

//                Array lines = File.ReadAllLines(filePath);
//                for(int i=0;i<row;i++)
//                {
//                    var file = File
//                    .ReadAllLines(filePath, Encoding.GetEncoding(932))
//                    .SkipWhile(line => string.IsNullOrWhiteSpace(line)) // To be on the safe side
//                    .Select((line, index) => index == i // do we have header? 
//                        ? line.Replace(",\"\"","") // replace '_' with ' '
//                        : line)                  // keep lines as they are 
//                    .ToList();                  // Materialization, since we write into the same file

//                    File.WriteAllLines(filePath, file, Encoding.GetEncoding(932));
//                }


//                using (TextFieldParser csvReader = new TextFieldParser(filePath, Encoding.GetEncoding(932), true))
//                {
//                    csvReader.SetDelimiters(new string[] { "," });
//                    csvReader.HasFieldsEnclosedInQuotes = true;
//                    //read column names
//                    string[] colFields = csvReader.ReadFields();
//                    char c = 'A';

//                    //for (int i = 0; i < colFields.Length; i++)
//                    //{
//                    //    colFields[i] = colFields[i].Replace("\"", "");
//                    //}

//                    //if (COL_COUNT != 0 && colFields.Length != COL_COUNT)
//                    //{
//                    //    int ct = colFields.Count();
//                    //    while (colFields.Count() > ct - 4)
//                    //    {
//                    //        colFields = colFields.Take(colFields.Count() - 1).ToArray();
//                    //    }


//                    //}

//                    foreach (string column in colFields)
//                    {
//                        if (firstColHeader.Equals("1")) //first row is column name
//                        {
//                            if (!csvData.Columns.Contains(column))
//                            {
//                                DataColumn datacolumn = new DataColumn(column);
//                                datacolumn.AllowDBNull = true;
//                                csvData.Columns.Add(datacolumn);
//                            }
//                            else
//                            {
//                                DataColumn datacolumn = new DataColumn(column + "_" + count++);
//                                datacolumn.AllowDBNull = true;
//                                csvData.Columns.Add(datacolumn);
//                            }
//                        }
//                        else//2
//                        {
//                            //if(!csvData.Columns.Contains("\"\"") || !csvData.Columns.Contains(""))
//                            //{
//                                DataColumn datacolumn = new DataColumn(c++.ToString());
//                                datacolumn.AllowDBNull = true;
//                                csvData.Columns.Add(datacolumn);
//                           // }

//                        }
//                    }

//                    if(firstColHeader.Equals("2")) // first row is data
//                        csvData.Rows.Add(colFields);//add first row as data row

//                    while (!csvReader.EndOfData)
//                    {
//                        string[] fieldData = csvReader.ReadFields();

//                        while (fieldData.Count() > colFields.Count())
//                        {
//                            fieldData = fieldData.Take(fieldData.Count() - 1).ToArray();
//                        }

//                        while (fieldData.Count() < colFields.Count())
//                        {
//                            Array.Resize(ref fieldData, fieldData.Length + 1);
//                            fieldData[fieldData.Length - 1] = null;
//                        }

//                        //Making empty value as null
//                        for (int i = 0; i < fieldData.Length; i++)
//                        {
//                            if (fieldData[i] == "")
//                            {
//                                fieldData[i] = null;
//                            }
//                        }
//                        csvData.Rows.Add(fieldData);

//                    }
//                }
//            }
//            catch (Exception ex)
//            {
//            }
//            return csvData;
//        }

//        public static int CountStringOccurrences(string[] arr, String item)
//        {
//            int i; int j = 0;
//            for (i = 0; i < arr.Count(); i++)
//            {
//                if (!string.IsNullOrEmpty(arr[i]))
//                {
//                    if (arr[i].Contains(item))
//                    {
//                        j++;
//                    }
//                }
//            }
//            return j;
//        }

//        public static DataTable CsvToDataTable(string FilePath)
//        {
//            DataTable csvData = new DataTable();
//            try
//            {
//                using (TextFieldParser csvReader = new TextFieldParser(FilePath, Encoding.GetEncoding(932)))
//                {
//                    csvReader.SetDelimiters(new string[] { "," });

//                    csvReader.HasFieldsEnclosedInQuotes = true;

//                    string[] colFields = csvReader.ReadFields();

//                    foreach (string column in colFields)
//                    {

//                        DataColumn datecolumn = new DataColumn(column);

//                        datecolumn.AllowDBNull = true;

//                        csvData.Columns.Add(datecolumn);
//                    }

//                    while (!csvReader.EndOfData)
//                    {
//                        string[] fieldData = csvReader.ReadFields();

//                        //Making empty value as null

//                        for (int i = 0; i < fieldData.Length; i++)
//                        {
//                            if (fieldData[i] == "")
//                            {
//                                fieldData[i] = null;
//                            }
//                        }
//                        csvData.Rows.Add(fieldData);
//                    }
//                }
//            }

//            catch (Exception)
//            {

//            }
//            return csvData;
//        }

//        private void btnOK_Click(object sender, EventArgs e)
//        {
//            this.Close();
//        }
//    }
//}
