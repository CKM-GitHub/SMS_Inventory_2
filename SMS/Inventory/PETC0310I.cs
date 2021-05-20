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
using SMS_DL;
using System.IO;
using System.Diagnostics;
using System.Globalization;

namespace SMS.Inventory
{
    public partial class frm0310I : SMS_Base
    {

        CommonFunctions cmf;
        Login_Info loginInfo;
        PETC0310I_DL pdl = new PETC0310I_DL();
        PETC0310I_BL p10 = new PETC0310I_BL();
        DataTable dtmain = new DataTable();
        public frm0310I(Login_Info loginInfo)
            : base(loginInfo)
        {
            InitializeComponent();
            this.loginInfo = loginInfo;
        }
        private void PETC0310I_Load(object sender, EventArgs e)
        {
            lblMode.Visible = false;
            FormName = "Tennic仕訳連携";//手数料仕訳作成(Tennic)
            F12_Text = "出力(F12)";
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
            cmf = new CommonFunctions();

            BindInitial();
        }
        public void BindInitial()
        {
            var dt = p10.T_SmileHistory();
            if (dt.Rows.Count > 0)
            {
                txtProcessNo.Text = dt.Rows[0]["Smile01"].ToString();
                txtDate.Text = Convert.ToDateTime(dt.Rows[0]["Smile03"].ToString()).ToString("yyyy/MM/dd");
                txtSlipNo.Text = dt.Rows[0]["Smile04"].ToString();
            }
            else
            {
                txtProcessNo.Text = txtDate.Text = txtSlipNo.Text = txtInput.Text = "";
            }

        }
        public override void functionClick(int btnIndex)
        {
            switch (btnIndex)
            {
                case 1:
                    break;

                case 11:
                 //   F11();
                    break;

                case 12:
                    F12();
                    break;
            }
        }
        protected bool ErrorCheck()
        {
            if (string.IsNullOrEmpty(txtInput.Text))
            {
                var msg = cmf.DSP_MSG("E102", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                button1.Focus();
                return false;
            }
            if (!string.IsNullOrEmpty(txtInput.Text) )
            {
                var ext = Path.GetExtension(txtInput.Text).Contains("txt");
                if (!ext)
                {
                    var msg = cmf.DSP_MSG("E102", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                    button1.Focus();
                    return false;
                }
            }
            return true;
        }
        protected void F12()
        {
            DialogResult dialogResult = MessageBox.Show("在庫データCSVを作成しますか？", ""  , MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
            if (dialogResult == DialogResult.OK)
            {
                if (!ErrorCheck())
                {
                    return;
                }
                Base_BL baseBL = new Base_BL();
                Cursor = Cursors.WaitCursor;
                if (dtmain == null)
                {
                    MessageBox.Show("Error in reading Files ... ");
                    return;
                }
                var de = p10.T_SmileInsert(loginInfo.OperatorCode, baseBL.DataTableToXml(dtmain));
                if(de)
                {
                    if (String.IsNullOrEmpty(txtProcessNo.Text) && string.IsNullOrWhiteSpace(txtSlipNo.Text)) // first time to input
                    {
                       //
                    }
                    else
                    {
                        var dtsmile = p10.GetTSmile();
                        if (dtsmile.Rows.Count > 0)
                        {
                           if (Convert.ToInt64(txtProcessNo.Text) >= Convert.ToInt64(dtsmile.Rows[0]["ProcessNo"].ToString())  || Convert.ToDateTime(txtDate.Text) >= Convert.ToDateTime(dtsmile.Rows[0]["Date"].ToString()))
                            {
                              var msg=  cmf.DSP_MSG("Q328", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                                if (msg != "1")
                                {
                                    Cursor = Cursors.Default; ;
                                    txtDate.Focus();
                                    return;
                                }
                            }
                        }
                    }

                    var ds = p10.T_SmileHistoryInsert_Output();
                    ConmmaSplit(ds);
                    var path = @"C:\SMS\PETC0310I\";
                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }
                    try
                    {
                        var dtnow = DateTime.Now.ToString("yyyyMMdd HHmmss");
                        var fname = dtnow.Replace(" ", "_") + "_" + loginInfo.OperatorCode + "_" + Path.GetFileNameWithoutExtension(txtInput.Text) + ".CSV";
                        if (ds.Rows.Count > 0)
                        {
                            if (File.Exists(path + "\\" + fname))
                            {
                                File.Delete(path + "\\" + fname);
                            }
                            ToCSV(ds, path + "\\" + fname);// For tennic
                        }
                        cmf.DSP_MSG("I201", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                        Init();
                        Process.Start(Path.GetDirectoryName(path));
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                    
                }
                Cursor = Cursors.Arrow;
                //do something
            }
            else if (dialogResult == DialogResult.No)
            {
                return;
                //do something else
            }

        }
        protected void Init()
        {
            BindInitial();
            txtInput.Clear();
            dtmain = null;
        }
        protected void ConmmaSplit(DataTable dt)
        {
            foreach (DataRow dr in dt.Rows)
            {
                //取引日付
                dr["取引日"] = dr["取引日"].ToString().Split(' ').First();
                dr["借方金額(円)"] = Convert.ToDecimal(dr["借方金額(円)"].ToString()).ToString("#,##0");
                dr["借方税額"] = Convert.ToDecimal(dr["借方税額"].ToString()).ToString("#,##0");
                dr["貸方金額(円)"] = Convert.ToDecimal(dr["貸方金額(円)"].ToString()).ToString("#,##0");
                dr["貸方税額"] = Convert.ToDecimal(dr["貸方税額"].ToString()).ToString("#,##0");
            }

        }
        public static void ToCSV(DataTable dtDataTable, string strFilePath)
        {
            //取引No		
            int o = 0;
            var PrevCol = "";
            foreach (DataRow dr in dtDataTable.Rows)
            {
                if (PrevCol != dr["Number"].ToString())
                {
                    o++;
                }
                dr["取引No"] = o.ToString();
                PrevCol = dr["Number"].ToString();
            }
            dtDataTable.Columns.Remove("Number");
            dtDataTable.AcceptChanges();
            StreamWriter sw = new StreamWriter(strFilePath, false, Encoding.GetEncoding(932));
            //headers  

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

        private void button1_Click(object sender, EventArgs e)
        {
            DataTable dtPath = new DataTable();

            //dtPath = PETC0309IBL.GetFolder3i();
            string folderPath = @"C:\SMS\PETC0310I";
            //folderPath = dtPath.Rows[0]["Char2"].ToString();
            if (!string.IsNullOrWhiteSpace(folderPath))
            {
                if (!Directory.Exists(folderPath))
                {
                    Directory.CreateDirectory(folderPath);
                }
                OpenFileDialog op = new OpenFileDialog();
                op.Multiselect = false;
                op.InitialDirectory = folderPath;
                op.RestoreDirectory = true;
                op.Filter = "Tennic files |*.txt";
                if (op.ShowDialog() == DialogResult.OK)
                {
                    if (op.SafeFileNames.Count() == 1)
                    {
                        txtInput.Text = op.FileName;
                    }
                }
                else
                    return;
                dtmain = TXTToDataTable(txtInput.Text, 0);

            }
        }

        private void smS_Button1_Click(object sender, EventArgs e)
        {
            F12();
        }
        public DataTable TXTToDataTable(string filePath, int numberOfColumns = 11)    // for error Check E137 
        {
            DataTable tbl = new DataTable();
            string[] lines = System.IO.File.ReadAllLines(filePath,  Encoding.UTF8);
            //Check First is Not 
            int cout = 0;
            foreach (string line in lines)
            {
                cout++;
                var cols = line.Split(',');
                if (cout == 1)
                {
                    if (cols[0].Replace("\"", string.Empty) != "処理連番")  // Return Error if not this text
                    {
                        cmf.DSP_MSG("E137", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                        return null;
                    }
                    int C = 0; 
                    foreach (string str in cols)
                    {
                        C++;
                        tbl.Columns.Add(C.ToString().Replace("\"", ""));
                    }
                }
                else
                {
                     cols = (new System.Text.RegularExpressions.Regex("\\\"(.*?)\\\"")).Replace(line, m => m.Value.Replace(",", string.Empty)).Split(',');
                    //var cols = (new System.Text.RegularExpressions.Regex("\\\"(.*?)\\\"")).Replace(line, m => m.Value.Replace(",", string.Empty)).Split(',');
                    DataRow dr = tbl.NewRow();
                    for (int cIndex = 0; cIndex < cols.Count(); cIndex++)
                    {
                        if (cIndex == 2 || cIndex == 87)
                        {
                            var str = cols[cIndex].ToString().Replace("\"", "");
                            dr[cIndex] = str.Substring(0, 4) + "-" + str.Substring(4, 2) + "-" + str.Substring(6, 2);
                        }
                        else
                        {
                            dr[cIndex] = cols[cIndex].ToString().Replace("\"", "");
                        }

                    }
                    tbl.Rows.Add(dr);
                }
            }

            return tbl;
        }
    }
}
