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
using SMS_BL;
using System.Data;
using System.IO;
using System.Diagnostics;
using ElencySolutions.CsvHelper;

namespace SMS.Inventory
{
    public partial class frmPETC0301I : SMS_Base
    {
        Login_Info loginInfo;
        PETC0301I_BL PETC0301I_bl;
        string startdate = string.Empty;
        string enddate = string.Empty;
        string ShopId = string.Empty;

        #region constructor
        public frmPETC0301I()
        {
            InitializeComponent();
        }

        public frmPETC0301I(Login_Info loginInfo)
            : base(loginInfo)
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
        /// 
        //public Tuple<string, string> GetRoles(string strGetRoles)
        //{
        //    DataTable dt = new System.Data.DataTable();
        //    var f = "-1";
        //    var s = "-1";
        //     foreach(DataRow dr in dt.Rows)
        //     {
        //         if (strGetRoles == "qwq")
        //         {
        //             f = "12";
        //             s = "212";
        //             break;
        //         }
        //     }
        //     return Tuple.Create(f, s);
        //}
        private void frmPETC0301I_Load(object sender, EventArgs e)
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

            FormName = "Yahoo 出荷番号出力";
            F12_Text = "出力(F12)";
            
            lblMode.Visible = false;
            BindCombo();
            txtStartDate.Text = DateTime.Now.ToString("yyyy/MM/dd");
            txtEndDate.Text = DateTime.Now.ToString("yyyy/MM/dd");
        }

        /// <summary>
        /// Bind Yanhoo Shops to ComboBox
        /// </summary>
        private void BindCombo() 
        {
            PETC0301I_bl = new PETC0301I_BL();
            DataTable dt = new DataTable();
            dt = PETC0301I_bl.YahooShop_Combo();

            cbo_YahooShop.DisplayMember = "Char1";
            cbo_YahooShop.ValueMember = "Key";
            cbo_YahooShop.DataSource = dt;
        }
       
        /// <summary>
        ///  override base form's virtual function
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
        /// Error Check for F12 Click
        /// </summary>
        /// <returns></returns>
        private bool ErrorCheck()
        {
            if (string.IsNullOrWhiteSpace(txtStartDate.Text) || txtStartDate.Text.ToString() == "0")  // Error Check for dateTextBox
            {
                DSP_MSG("E102", "名前", string.Empty, string.Empty, string.Empty, string.Empty);
                txtStartDate.Focus();
                return false;
            }
            else if (string.IsNullOrWhiteSpace(txtEndDate.Text) || txtEndDate.Text.ToString() == "0")
            {
                DSP_MSG("E102", "名前", string.Empty, string.Empty, string.Empty, string.Empty);
                txtEndDate.Focus();
                return false;
            }
            else if (!string.IsNullOrWhiteSpace(txtStartDate.Text) && !string.IsNullOrWhiteSpace(txtEndDate.Text))
            {
                if (Convert.ToDateTime(txtStartDate.Text) > Convert.ToDateTime(txtEndDate.Text))
                {
                    DSP_MSG("E104", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                    txtEndDate.Focus();
                    return false;
                }
            }
            return true;
        }

        /// <summary>
        /// Export CSV file
        /// </summary>
        private void F12()
        {
            if(ErrorCheck())
            {
                ShopId = cbo_YahooShop.SelectedValue.ToString();
                startdate = txtStartDate.Text.Replace("/", "-");
                enddate = txtEndDate.Text.Replace("/", "-");

                DataTable dtYahoo = new DataTable();
                PETC0301I_bl = new PETC0301I_BL();
                dtYahoo = PETC0301I_bl.PECT0301I_Select(ShopId, startdate, enddate);
                if (dtYahoo.Rows.Count > 0)
                {
                    PETC0301I_bl = new PETC0301I_BL();
                    DataTable dtFolder = new DataTable();
                    dtFolder = PETC0301I_bl.GetFolder(ShopId);
                    string Folderpath = dtFolder.Rows[0]["Char2"].ToString();
                    if (!string.IsNullOrWhiteSpace(Folderpath))
                    {
                        if (!Directory.Exists(Folderpath))
                        {
                            Directory.CreateDirectory(Folderpath);
                        }
                        #region CSV create and save
                        SaveFileDialog savedialog = new SaveFileDialog();
                        savedialog.Filter = "CSV|*.csv";
                        savedialog.Title = "Save";
                        savedialog.FileName = dtFolder.Rows[0]["Char3"].ToString();
                        savedialog.InitialDirectory = dtFolder.Rows[0]["Char2"].ToString();
                        savedialog.RestoreDirectory = true;
                        if (savedialog.ShowDialog() == DialogResult.OK)
                        {
                            if (Path.GetExtension(savedialog.FileName).Contains("csv"))
                            {
                                CsvWriter csvwriter = new CsvWriter();
                                csvwriter.WriteCsv(dtYahoo, savedialog.FileName, Encoding.GetEncoding(932));
                            }
                            Process.Start(Path.GetDirectoryName(savedialog.FileName));
                        }
                        #endregion
                    }
                }
                else
                {
                    MessageBox.Show("There's no data to export!", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }
    }
}
