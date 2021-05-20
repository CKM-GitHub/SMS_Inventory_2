using Microsoft.VisualBasic.FileIO;
using SMS_BL;
using SMS_Entity;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SMS_BL;
using System.Diagnostics;
using ElencySolutions.CsvHelper;

namespace SMS.Inventory
{
    public partial class frmPETC0306I : SMS_Base
    {
        Login_Info loginInfo = new Login_Info();
        M_MultiPorpose_Entity mme = new M_MultiPorpose_Entity();
        L_Log_Entity LLogEntity;
        PSKS0103I_BL psks0103ibl = new PSKS0103I_BL();
        PETC0306I_BL petc0306ibl = new PETC0306I_BL();
        T_AmazonYoteiDate_Entity amazonyoteidate_data;
        string folderPath = string.Empty;
        string fileName = string.Empty;
        public frmPETC0306I()
        {
            InitializeComponent();
        }
        public frmPETC0306I(Login_Info loginInfo)
            : base(loginInfo)
        {
            InitializeComponent();
            this.loginInfo = loginInfo;
        }
        private void frmPETC0306I_Load(object sender, EventArgs e)
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
            FunctionButtonDisabled(11);

            lblMode.Visible = false;
            FormName = "Amazonデータ作成（出荷予定日）";

            F12_Text = "取込＆出力(F12)";

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

                case 12:
                    F12();
                    break;
            }
        }
        /// <summary>
        /// folder open funcion
        /// </summary>
        /// <param name="sender,e"></param>
        private void btnForderOpen_Click(object sender, EventArgs e)
        {
            mme.fields = "Char1,Char2";
            mme.tableName = "M_MultiPorpose";
            mme.condition = "Where ID=116 And [Key]='1'";
            DataTable dt = psks0103ibl.M_MultiPorpose_DynamicSelect(mme);
            if (dt.Rows.Count > 0)
            {
                folderPath = dt.Rows[0]["Char2"].ToString();
                fileName = dt.Rows[0]["Char1"].ToString();
            }
            OpenFileDialog dlog = new OpenFileDialog();
            dlog.Filter = "TXT|*.txt";
            dlog.InitialDirectory = folderPath;
            if (!string.IsNullOrWhiteSpace(folderPath))
            {

                if (!Directory.Exists(folderPath))
                {
                    try
                    {
                        Directory.CreateDirectory(folderPath);
                    }
                    catch
                    {
                        string newPath = @"C:\";
                        dlog.InitialDirectory = newPath;
                        Directory.CreateDirectory(newPath);
                    }
                }
            }
            if (dlog.ShowDialog() == DialogResult.OK)
            {
                txtfilepath.Text = dlog.FileName;
                txtfilepath.BackColor = Color.White;

                // update new path in multiporpose table 
                string folder = txtfilepath.Text.Remove(txtfilepath.Text.LastIndexOf('\\'));
                mme.fields = Path.GetFileName(txtfilepath.Text) + "," + folder;
                mme.tableName = "M_MultiPorpose";
                mme.condition = " Where ID=116 And [Key]='1'";

                var data = psks0103ibl.M_MultiPorpose_DynamicUpdateData(mme);
            }      

        }

        // <summary>
        ///データを入力、更新する、CSV出力
        /// </summary>
        /// <param></param>
        /// <returns></returns>
        public void F12()
        {
            if (ErrorCheck())
            {
                if (DSP_MSG("Q101", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty).Equals("1"))
                {
                    bool valid = true;
                    DataTable dtImport = new DataTable();
                    int rowcount=0;
                    dtImport = CsvToDataTable(txtfilepath.Text);
                    amazonyoteidate_data = new T_AmazonYoteiDate_Entity();
                    LLogEntity = new L_Log_Entity();
                     amazonyoteidate_data.OperatorCD = loginInfo.OperatorCode;

                    if (IsEmptyGrid(dtImport))
                    {
                        DSP_MSG("E137", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                        txtfilepath.Focus();
                        valid = false;
                    }
                    else {
                        string aa=dtImport.Rows[0]["A"].ToString();
                        //if (!(dtImport.Rows[0]["A"].ToString().Equals("order-id")) && !(dtImport.Rows[0]["B"].ToString().Equals("order-item-id")) && !(dtImport.Rows[0]["C"].ToString().Equals("purchase-date")))
                        if (!((dtImport.Rows[0]["A"].ToString().Equals("order-id")) && (dtImport.Rows[0]["B"].ToString().Equals("order-item-id")) && (dtImport.Rows[0]["C"].ToString().Equals("purchase-date"))))
                    {
                        DSP_MSG("E137", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                        txtfilepath.Focus();
                        valid = false;
                    }
                    }

                    if (valid)
                    {
                        if (petc0306ibl.PETC0306I_DELETE(amazonyoteidate_data))
                        {   
                            for (int i = 1; i < dtImport.Rows.Count;i++ )
                            {
                                rowcount = rowcount + 1;
                                amazonyoteidate_data.OperatorCD = loginInfo.OperatorCode;
                                amazonyoteidate_data.SEQ = rowcount.ToString();
                                amazonyoteidate_data.amazon_order_id = dtImport.Rows[i]["A"].ToString();
                                amazonyoteidate_data.sku = dtImport.Rows[i]["H"].ToString();
                                //insert data into database from csv
                                petc0306ibl.PETC0306I_Insert(amazonyoteidate_data);
                                //update estimated date data into databse
                                petc0306ibl.PETC0306I_Update(amazonyoteidate_data);
                                
                            }
                            Export_CSV(petc0306ibl.PETC0306I_Select(amazonyoteidate_data));

                        }
                    }
                    
               }

            }

        }
        // <summary>
        /// Export CSV
        /// </summary>
        /// <param name="dtdgv"></param>
        /// <returns></returns>
        public void Export_CSV(DataTable dtdgv){
            
            mme.fields = "Char3,Char2";
            mme.tableName = "M_MultiPorpose";
            mme.condition = "Where ID=116 And [Key]='1'";
            DataTable dt = psks0103ibl.M_MultiPorpose_DynamicSelect(mme);

            string FolderPath = dt.Rows[0]["Char2"].ToString();
            //string FolderPath = folderPath;
            string filename = dt.Rows[0]["Char3"].ToString();
           
            if (dtdgv.Rows.Count > 0)
            {
              
                if (!string.IsNullOrWhiteSpace(FolderPath))
                {
                    if (!Directory.Exists(FolderPath))
                    {
                        FolderPath = "C:";
                        Directory.CreateDirectory(FolderPath);
                    }
                    SaveFileDialog savedialog = new SaveFileDialog();
                    savedialog.Filter = "TXT|*.txt";
                    savedialog.Title = "Save";
                    savedialog.FileName = filename;
                    savedialog.InitialDirectory = FolderPath;
                    savedialog.RestoreDirectory = true;

                    if (savedialog.ShowDialog() == DialogResult.OK)
                    {
                        if (Path.GetExtension(savedialog.FileName).Contains("txt"))
                        {
                            CsvWriter csvwriter = new CsvWriter(); 
                            csvwriter.WriteCsv(dtdgv, savedialog.FileName, Encoding.GetEncoding(932));
                            //convert comma sperator to tab sperator 
                            var csv = File.ReadAllLines(savedialog.FileName).Select(line => line.Replace(",", "\t"));
                            File.WriteAllLines(savedialog.FileName, csv);
 
                           
                        }
                        Process.Start(Path.GetDirectoryName(savedialog.FileName));
                    }
                }

            }
            else
            {
               MessageBox.Show("There's no data to export!", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        // <summary>
        /// validate data when No data in datatable
        /// </summary>
        /// <param name="datatable"></param>
        /// <returns></returns>
        public bool IsEmptyGrid(DataTable datatable)
        {
             for(int i=0;i<datatable.Rows.Count;i++)
                {
               for(int j=0;j<datatable.Columns.Count;j++)
                {
                if(string.IsNullOrEmpty(datatable.Rows[i][j].ToString()))
                {
                return false;
                }
             }
             }
            return true;
        }   

        // <summary>
        /// Move CSV file into Datatable
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static DataTable CsvToDataTable(string FilePath)
        {
            DataTable csvData = new DataTable();
            try
            {
                string[] colFields = "A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,AA,AB,AC,AD,AE,AF,AG,AH,AI,AJ,AK,AL,AM,AN,AO,AP,AQ,AR,AS,AT,AU,AV".Split(',');

                foreach (string column in colFields)
                {

                    DataColumn datecolumn = new DataColumn(column);

                    datecolumn.AllowDBNull = true;

                    csvData.Columns.Add(datecolumn);
                }
                using (TextFieldParser csvReader = new TextFieldParser(FilePath, Encoding.GetEncoding(932)))
                {
                    csvReader.SetDelimiters(new string[] { "\t" });

                    csvReader.HasFieldsEnclosedInQuotes = true;

                    while (!csvReader.EndOfData)
                    {
                        string[] fields = csvReader.ReadFields();

                        while (fields.Count() > colFields.Count())
                        {
                            fields = fields.Take(fields.Count() - 1).ToArray();
                        }

                        while (fields.Count() < colFields.Count())
                        {
                            Array.Resize(ref fields, fields.Length + 1);
                            fields[fields.Length - 1] = null;
                        }

                        //Making empty value as null
                        for (int i = 0; i < fields.Length; i++)
                        {
                            if (fields[i] == "")
                            {
                                fields[i] = null;
                            }
                        }
                        csvData.Rows.Add(fields);
                    }
                }
            }
            catch (Exception)
            {

            }
            return csvData;
        }

        /// <summary>
        /// Error Check Before Import
        /// </summary>
        /// <returns></returns>
        private bool ErrorCheck()
        {
            if (!string.IsNullOrWhiteSpace(txtfilepath.Text))
            {
                string folderPath = txtfilepath.Text.Remove(txtfilepath.Text.LastIndexOf('\\'));
                string filePath = txtfilepath.Text;
                if (!Directory.Exists(folderPath))
                {
                    Directory.CreateDirectory(folderPath);
                }

                if (File.Exists(filePath))
                {
                    if (!(Have_opened_Excel(filePath)))
                    {
                        MessageBox.Show("Some of the imported files are being used by another process. please close those files to be able to import", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return false;
                    }
                }
                else
                {
                    MessageBox.Show("File doesn't exit in this path ", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
                return true;
            }
            return false;
        }

        /// <summary>
        /// check folder is open or not
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public bool Have_opened_Excel(string filePath)
        {
            try
            {
                var f = File.OpenRead(filePath);
                f.Close();
            }
            catch
            {
                return false;
            }

            return true;
        }

    }
}
