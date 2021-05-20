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
using SMS.CustomControls;
using SMS_BL;
using Microsoft.VisualBasic.FileIO;

namespace SMS
{
    public partial class SMS_Base : Form
    {
        CommonFunctions cf = new CommonFunctions();
        Base_BL basebl = new Base_BL();
        public bool flag = false; // 30/11/2018
        
        public SMS_Base()
        {
            InitializeComponent();
        }

        public string FormMode
        {
            set { lblMode.Text = value; }
            get { return lblMode.Text; }
        }

        public string FormName
        {
            set { lblFormName.Text = value; }
            get { return lblFormName.Text; }
        }

        public string Operator
        {
            set { lblSOperator.Text = value; }
            get { return lblSOperator.Text; }
        }

        public SMS_Base(Login_Info loginInfo)
        {
            InitializeComponent();
            FormMode = loginInfo.StartMode;
            Operator = loginInfo.OperatorName;
            
            lblSProcessingDate.Text = DateTime.Today.ToString("yyyy/MM/dd");
        }

        public string F1_Text
        {
            get { return btnF1.Text; }
            set { btnF1.Text = value; }
        }

        public string F5_Text
        {
            get { return btnF5.Text; }
            set { btnF5.Text = value; }
        }

        public string F6_Text
        {
            get { return btnF6.Text; }
            set{btnF6.Text = value;}
        }

        public string F7_Text
        {
            get { return btnF7.Text; }
            set { btnF7.Text = value; }
        }

        public string F8_Text
        {
            get { return btnF8.Text; }
            set { btnF8.Text = value; }
        }

        public string F9_Text
        {
            get { return btnF9.Text; }
            set { btnF9.Text = value; }
        }

        public string F10_Text
        {
            get { return btnF10.Text; }
            set { btnF10.Text = value; }
        }

        public string F11_Text
        {
            get { return btnF11.Text; }
            set { btnF11.Text = value; }
        }

        public string F12_Text
        {
            get { return btnF12.Text; }
            set { btnF12.Text = value; }
        }

        private void SetDesignerFunction()
        {
            this.KeyPreview = true;
            this.KeyDown += frmKeyDown;

            btnF1.Click += btn_Click;
            btnF2.Click += btn_Click;
            btnF3.Click += btn_Click;
            btnF4.Click += btn_Click;
            btnF5.Click += btn_Click;
            btnF6.Click += btn_Click;
            btnF7.Click += btn_Click;
            btnF8.Click += btn_Click;
            btnF9.Click += btn_Click;
            btnF10.Click += btn_Click;
            btnF11.Click += btn_Click;
            btnF12.Click += btn_Click;
        }

        private void btn_Click(object sender, EventArgs e)
        {
            Button btn = (Button)sender;
            if (btn.Cursor == Cursors.Hand)
            {
                int index = Convert.ToInt32(btn.Name.Substring(4));
                Base_functionClick(index);
            }
        }

        public void frmKeyDown(object sender, KeyEventArgs e)
        
        {
            if (e.KeyCode == Keys.Enter)
            {
                SendKeys.Send("{TAB}");
            }
            //else if (e.KeyCode == Keys.F10)
            //{
            //    SendKeys.Send("{F10}");
            //}
            else if ((e.KeyCode == Keys.F1 && (btnF1.Cursor == Cursors.Hand)) ||
                    (e.KeyCode == Keys.F2 && (btnF2.Cursor == Cursors.Hand)) ||
                    (e.KeyCode == Keys.F3 && (btnF3.Cursor == Cursors.Hand)) ||
                    (e.KeyCode == Keys.F4 && (btnF4.Cursor == Cursors.Hand)) ||
                    (e.KeyCode == Keys.F5 && (btnF5.Cursor == Cursors.Hand)) ||
                    (e.KeyCode == Keys.F6 && (btnF6.Cursor == Cursors.Hand)) ||
                    (e.KeyCode == Keys.F7 && (btnF7.Cursor == Cursors.Hand)) ||
                    (e.KeyCode == Keys.F8 && (btnF8.Cursor == Cursors.Hand)) ||
                    (e.KeyCode == Keys.F9 && (btnF9.Cursor == Cursors.Hand)) ||
                    (e.KeyCode == Keys.F10 && (btnF10.Cursor == Cursors.Hand)) ||
                    (e.KeyCode == Keys.F11 && (btnF11.Cursor == Cursors.Hand)) ||
                    (e.KeyCode == Keys.F12 && (btnF12.Cursor == Cursors.Hand)))
            {
                Base_functionClick(Convert.ToInt32(e.KeyCode.ToString().Substring(1)));
            }

        }

        protected void FunctionButtonDisabled(int no)
        {
            Button btn = (no == 1) ? btnF1 : (no == 2) ? btnF2 : (no == 3) ? btnF3 : (no == 4) ? btnF4 : (no == 5) ? btnF5 : (no == 6) ? btnF6 : (no == 7) ? btnF7 : (no == 8) ? btnF8 : (no == 9) ? btnF9 : (no == 10) ? btnF10 :(no ==11)?btnF11: btnF12;
            btn.ForeColor = btn.BackColor;
            btn.Cursor = Cursors.No;
        }

        protected void FunctionButtonEnabled(int no)
        {
            Button btn = (no == 1) ? btnF1 : (no == 2) ? btnF2 : (no == 3) ? btnF3 : (no == 4) ? btnF4 : (no == 5) ? btnF5 : (no == 6) ? btnF6 : (no == 7) ? btnF7 : (no == 8) ? btnF8 : (no == 9) ? btnF9 : (no == 10) ? btnF10 : (no == 11) ? btnF11 : btnF12;
            btn.ForeColor = Color.Black;
            btn.Cursor = Cursors.Hand;
        }

        protected void Base_functionClick(int index)
        {
            if (index == 1) // 30/11/2018
            {
                if (BaseFunctionClick(btnF1.Text))
                {
                    flag = true;
                }
            }
            else if (index == 2) 
            {
                if (BaseFunctionClick(btnF2.Text))
                {
                    flag = true;
                }
            }
            else if (index == 3) 
            {
                if (BaseFunctionClick(btnF3.Text))
                {
                    flag = true;
                }
            }
            else if (index == 4) 
            {
                if (BaseFunctionClick(btnF4.Text))
                {
                    flag = true;
                }
            }
            else if (index == 5) 
            {
                if (BaseFunctionClick(btnF5.Text))
                {
                    flag = true;
                }
            }
            else if (index == 6) 
            {
                if (BaseFunctionClick(btnF6.Text))
                {
                    flag = true;
                }
            }
            else if (index == 7) 
            {
                if (BaseFunctionClick(btnF7.Text))
                {
                    flag = true;
                }
            }
            else if (index == 8) 
            {
                if (BaseFunctionClick(btnF8.Text))
                {
                    flag = true;
                }
            }
            else if (index == 9) 
            {
                if (BaseFunctionClick(btnF9.Text))
                {
                    flag = true;
                }
            }
            else if (index == 10) 
            {
                if (BaseFunctionClick(btnF10.Text))
                {
                    flag = true;
                }
            }
            else if (index == 11) 
            {
                if (BaseFunctionClick(btnF11.Text))
                {
                    flag = true;
                }
            }
            else if (index == 12) 
            {
                if (BaseFunctionClick(btnF12.Text))
                {
                    flag = true;
                }
            }
            //if (BaseFunctionClick(index))
            if (flag == true)
            {
                functionClick(index);
            }
        }

        public virtual void functionClick(int index) { }

        //private bool BaseFunctionClick(int index)
        private bool BaseFunctionClick(string btnText)
        {
            //if(btnText)
            {
                if (btnText.Contains("戻る"))
                {
                    this.Close();
                }
                else if (btnText.Contains("終了"))
                //switch (index)
                {
                    //case 1:
                    if (DSP_MSG("Q003", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty).Equals("1"))
                    {
                        this.Close();
                    }
                    else
                        return false;
                    //break;
                }
                //case 2:
                else if (btnText.Contains("新規"))
                {
                    if (DSP_MSG("Q005", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty).Equals("1"))
                    {
                        NewMode();
                    }
                    else return false;
                    //    break;
                }
                //case 10: zaikoedit
                else if (btnText.Contains("在庫変更"))
                {
                    return true;
                }
                //case 3:
                else if (btnText.Contains("変更"))
                {
                    if (DSP_MSG("Q005", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty).Equals("1"))
                    {
                        FormMode = "修正";
                        lblMode.BackColor = Color.Yellow;
                        FunctionButtonEnabled(1);
                        FunctionButtonEnabled(2);
                        FunctionButtonEnabled(4);
                        FunctionButtonEnabled(5);
                        FunctionButtonEnabled(12);

                        FunctionButtonDisabled(3);
                        FunctionButtonDisabled(9);
                        FunctionButtonDisabled(11);
                    }
                    else return false;
                    //    break;
                }
                //case 4:
                else if (btnText.Contains("削除"))
                {
                    if (FormName == "在庫状況管理登録")
                    {
                        if (DSP_MSG("Q102", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty).Equals("1"))
                        {
                            FunctionButtonDisabled(9);
                        }
                        else
                        {
                            flag = false;
                            return false;
                        }
                    }
                    else
                    {
                        if (DSP_MSG("Q005", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty).Equals("1"))
                        {
                            FormMode = "削除";
                            lblMode.BackColor = Color.Red;
                            FunctionButtonEnabled(1);
                            FunctionButtonEnabled(2);
                            FunctionButtonEnabled(3);
                            FunctionButtonEnabled(5);
                            FunctionButtonEnabled(11);
                            FunctionButtonEnabled(12);

                            FunctionButtonDisabled(4);
                            FunctionButtonDisabled(9);
                        }
                        else
                            return false;
                    }
                    //    break;
                }
                //case 5:
                else if (btnText.Contains("照会"))
                {
                    if (DSP_MSG("Q005", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty).Equals("1"))
                    {
                        FormMode = "照会";
                        lblMode.BackColor = Color.DeepSkyBlue;
                        FunctionButtonEnabled(1);
                        FunctionButtonEnabled(2);
                        FunctionButtonEnabled(3);
                        FunctionButtonEnabled(4);

                        FunctionButtonDisabled(5);
                        FunctionButtonDisabled(9);
                        FunctionButtonDisabled(11);
                        FunctionButtonDisabled(12);
                    }
                    else
                        return false;
                    //    break;
                }
                else if (btnText.Contains("確定") || btnText.Contains("取込・登録") || btnText.Contains("登録") || btnText.Contains("取込・登録"))
                {
                    return true;
                }
                //else if (btnText.Contains("取込"))
                //{
                //    if (!DSP_MSG("Q101", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty).Equals("1"))
                //        return false;
                //}
                //else if (btnText.Contains("出力"))
                //{
                //    if (!DSP_MSG("Q203", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty).Equals("1"))
                //        return false;
                //}

                //case 6:
                else if (btnText.Contains("ｷｬﾝｾﾙ"))
                {
                    if (!DSP_MSG("Q004", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty).Equals("1"))
                        return false;
                    //break;
                }
                else if (btnText.Contains("リスト印刷"))
                {
                    return true;
                }
                else if (btnText.Contains("印刷"))
                {
                    return true;
                }              
            }
            return true;
        }

        protected void NewMode()
        {
            FormMode = "新規";
            lblMode.BackColor = Color.Orange;
            FunctionButtonEnabled(1);
            FunctionButtonEnabled(3);
            FunctionButtonEnabled(4);
            FunctionButtonEnabled(5);
            FunctionButtonEnabled(11);
            FunctionButtonEnabled(12);

            FunctionButtonDisabled(2);
            FunctionButtonDisabled(9);
            FunctionButtonDisabled(11);
        }

        public string DSP_MSG(string messageID, string messageParam1, string messageParam2, string messageParam3, string messageParam4, string messageParam5)
        {
            return cf.DSP_MSG(messageID, messageParam1, messageParam2, messageParam3, messageParam4, messageParam5);
        }

        private void SMS_Base_Load(object sender, EventArgs e)
        {
            SetDesignerFunction();
        }

        protected bool checkNullOrWhiteSpace(Dictionary<Control, string> dic)
        {
            foreach (KeyValuePair<Control, string> kvp in dic)
            {
                if (kvp.Key is TextBox)
                {
                    TextBox txt = kvp.Key as TextBox;
                    if (string.IsNullOrWhiteSpace(txt.Text))
                    {
                        DSP_MSG("E102", kvp.Value, string.Empty, string.Empty, string.Empty, string.Empty);
                        txt.Focus();
                        return false;
                    }
                }
                else if (kvp.Key is ComboBox)
                {
                    ComboBox cbo = kvp.Key as ComboBox;
                    if (cbo.SelectedIndex <= 0)
                    {
                        DSP_MSG("E102", kvp.Value, string.Empty, string.Empty, string.Empty, string.Empty);
                        cbo.Focus();
                        return false;
                    }
                }
            }
            return true;
        }

        protected void Clear(Panel p)
        {
            IEnumerable<Control> c = GetAllControls(p);
            foreach (Control ctrl in c)
            {
                if (ctrl is TextBox || ctrl is TextBox || ctrl is RichTextBox)
                {
                    ctrl.Text = string.Empty;
                }
                else if (ctrl is ComboBox)
                {
                    ((ComboBox)ctrl).SelectedIndex = -1;
                }
                else if (ctrl is CheckBox)
                {
                    ((CheckBox)ctrl).Checked = false;
                }
                else if (ctrl is DataGridView)
                {
                    DataGridView dgv = ctrl as DataGridView;
                    DataTable dt = dgv.DataSource as DataTable;
                    if (dt != null)
                        dt.Clear();
                    dgv.DataSource = dt;
                }
            }
        }

        private IEnumerable<Control> GetAllControls(Panel p)
        {
            return cf.GetAllControls(p);
        }

        protected bool NullCheckForGrid(DataGridView dg, DataTable dt, Dictionary<string, string> dic)//Check Null For Grid
        {
            int row = -1;
            foreach (DataRow dr in dt.Rows)
            {
                row++;
                if (!CheckAllColumnNull(dic, dr))
                {
                    foreach (KeyValuePair<string, string> kvp in dic)
                    {
                        string col1 = kvp.Key.Split(',')[0];//gridview column name
                        string col2 = kvp.Key.Split(',')[1];//datatable header name
                        if (string.IsNullOrWhiteSpace(dr[col2].ToString()))
                        {
                            DSP_MSG("E102", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                            dg.CurrentCell = dg.Rows[row].Cells[col1];
                            return false;
                        }
                    }
                }
            }

            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dg"></param>
        /// <param name="dt"></param>
        /// <param name="dic">fields,tableName,datatable's column name</param>
        /// <returns></returns>
        protected bool DataExistsCheckForGrid(DataGridView dg, DataTable dt, Dictionary<string, string> dic)
        {
            int row = -1;
            foreach (DataRow dr in dt.Rows)
            {
                row++;
                foreach (KeyValuePair<string, string> kvp in dic)
                {
                    string fields = kvp.Value.Split(',')[0];
                    string tableName = kvp.Value.Split(',')[1];
                    string colName = kvp.Key;
                    string condition = "Where " + fields + " = '" + dr[colName] + "'";

                    if (!string.IsNullOrWhiteSpace(dr[colName].ToString()))
                    {
                        DataTable dtResult = basebl.DynamicSelectData(fields, tableName, condition);
                        if (dtResult.Rows.Count <= 0)
                        {
                            DSP_MSG("E101", kvp.Key, string.Empty, string.Empty, string.Empty, string.Empty);
                            return false;
                        }
                    }
                }
            }

            return true;
        }

        private bool CheckAllColumnNull(Dictionary<string, string> dic, DataRow dr)//all column is blank
        {
            foreach (KeyValuePair<string, string> kvp in dic)
            {
                string col = kvp.Key.Split(',')[1];
                if (!string.IsNullOrWhiteSpace(dr[col].ToString()))
                    return false;
            }
            return true;
        }

        public static DataTable GetDataTabletFromCSVFile(string FilePath)
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

        public bool DateTimeNullCheck(DateTimeTextBox txt)
        {
            if (string.IsNullOrWhiteSpace(txt.Text.Replace(":", "").Trim()))
            {
                DSP_MSG("E102", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                txt.Focus();
                return false;
            }
            return true;
        }
    }
}
