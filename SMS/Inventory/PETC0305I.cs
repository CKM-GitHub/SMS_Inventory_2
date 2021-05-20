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
using SMS_Entity;
using SMS_BL;
using Microsoft.VisualBasic.FileIO;
using System.Collections;
using ElencySolutions.CsvHelper;
using ClosedXML.Excel;
using System.Diagnostics;
using DataTableToCSV;
using System.IO;
using System.Security.AccessControl;
using System.Security;
using System.Security.Permissions;
using System.Security.Principal;

namespace SMS.Inventory
{
    public partial class frmPETC0305I : SMS_Base
    {
        Login_Info loginInfo = new Login_Info();
        M_MultiPorpose_Entity mme = new M_MultiPorpose_Entity();
        PSKS0103I_BL psks0103ibl = new PSKS0103I_BL();
        string firstColHeader = string.Empty;

        string folderPath = string.Empty;
        string fileName = string.Empty;

        #region constructor
        public frmPETC0305I()
        {
            InitializeComponent();
        }

        public frmPETC0305I (Login_Info loginInfo):base(loginInfo )
        {
            InitializeComponent();
            this.loginInfo = loginInfo;
        }
        #endregion

        /// <summary>
        /// Form Load
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmPETC0305I_Load(object sender, EventArgs e)
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
            FormName = "楽天データ編集（キャンセル削除）";
            F12_Text = "取込＆出力(F12)";
            txtpath.BackColor = Color.White;
        }

        /// <summary>
        /// Override base form's virtual function 
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
        /// export Csv file
        /// </summary>
        public void F12()
        {
            if (ErrorCheck())
            {
                if (DSP_MSG("Q002", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty).Equals("1"))
                {
                    string Folderpath = txtpath.Text.Remove(txtpath.Text.LastIndexOf('\\'));

                    if (Path.GetExtension(txtpath.Text).Equals(".csv"))
                    {
                        DataTable dtImport = new DataTable();
                        dtImport = CsvToDataTable(txtpath.Text); 

                        dtImport.Select("ステータス='900'").ToList().ForEach(x => x.Delete());
                        dtImport.AcceptChanges();  

                        if (dtImport != null)
                        {
                            if (!string.IsNullOrWhiteSpace(Folderpath))
                            {
                                if (!Directory.Exists(Folderpath))
                                {
                                    string newPath = @"c:\";
                                    Directory.CreateDirectory(newPath);
                                   
                                }

                                DataTableToCSV.Program ptj = new DataTableToCSV.Program(); //datatable to csv
                                ptj.CSV(dtImport, txtpath.Text, true);
                            }
                        }         
              
                        //DSP_MSG("I201", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);

                        mme.fields = "Char1,Char2";
                        mme.tableName = "M_MultiPorpose";
                        mme.condition = "Where ID=113 And [Key]='1'";

                        DataTable dt = psks0103ibl.M_MultiPorpose_DynamicSelect(mme);
                        if (dt.Rows.Count > 0)
                        {
                            folderPath = dt.Rows[0]["Char1"].ToString();
                            fileName = dt.Rows[0]["Char2"].ToString();
                        }

                        OpenFileDialog dlg = new OpenFileDialog();
                        dlg.InitialDirectory = folderPath;
                        dlg.FileName = dt.Rows[0]["Char1"].ToString()+"\\"+fileName;
                        dlg.RestoreDirectory = true;
                        Process.Start(Path.GetDirectoryName(dlg.FileName));
                    }
                    else
                    {
                        MessageBox.Show("Your File need to be CSV Type.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    }
            }
        }

        /// <summary>
        /// Check error before Import
        /// </summary>
        /// <returns></returns>
        private bool ErrorCheck()
        {
            if (!string.IsNullOrWhiteSpace(txtpath.Text))
            {
                string folderPath = txtpath.Text.Remove(txtpath.Text.LastIndexOf('\\'));
                string filePath = txtpath.Text;
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
        /// Open floder
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnFolderOpen_Click(object sender, EventArgs e)
        {
            mme.fields = "Char1,Char2";
            mme.tableName = "M_MultiPorpose";
            mme.condition = "Where ID=113 And [Key]='1'";

            DataTable dt = psks0103ibl.M_MultiPorpose_DynamicSelect(mme);
            if (dt.Rows.Count > 0)
            {
                folderPath = dt.Rows[0]["Char1"].ToString();
                fileName = dt.Rows[0]["Char2"].ToString();
            }

            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "CSV|*.csv";
            dlg.InitialDirectory = folderPath;

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
                        dlg.InitialDirectory = newPath;
                        Directory.CreateDirectory(newPath);
                    }
                }
            }
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                txtpath.Text = dlg.FileName;
                txtpath.BackColor = Color.White;

                // update new path in multiporpose table 
                string file=txtpath.Text;
                string folder = txtpath.Text.Remove(txtpath.Text.LastIndexOf('\\'));
               // mme.fields = Path.GetFileName(txtpath.Text) + "," + folder;
                mme.fields = folder + "," + Path.GetFileName(txtpath.Text);
               // mme.fields = file + "," + Path.GetFileName(txtpath.Text);
                mme.tableName = "M_MultiPorpose";
                mme.condition = " Where ID=113 And [Key]='1'";

                var data = psks0103ibl.M_MultiPorpose_DynamicUpdateData(mme);
            }
        }
        /// <summary>
        /// Check the floder if open or not 
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

        /// <summary>
        /// Move Csv filde to DataTable
        /// </summary>
        /// <param name="FilePath"></param>
        /// <returns></returns>
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

    }
}
