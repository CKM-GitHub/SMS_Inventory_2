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
using ElencySolutions.CsvHelper;
using System.Diagnostics;

namespace SMS.Inventory
{
    public partial class frmPETC0307I : SMS_Base
    {
        Login_Info loginInfo;
        PETC0307I_BL petc0307_bl;
        string startdate = string.Empty;
        string enddate = string.Empty;
        string ShopId = string.Empty;


        public frmPETC0307I()
        {
            InitializeComponent();
        }

        public frmPETC0307I(Login_Info loginInfo)
            : base(loginInfo)
        {
            InitializeComponent();
            this.loginInfo = loginInfo;
        }

        private void frmPECT0307I_Load(object sender, EventArgs e)
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

            FormName = "楽天 出荷番号出力";
            F12_Text = "出力(F12)";

            lblMode.Visible = false;
            BindCombo();
            txtStartDate.Text = DateTime.Now.ToString("yyyy/MM/dd");
            txtEndDate .Text = DateTime.Now.ToString ("yyyy/MM/dd");

        }

        private void BindCombo()
        {
            petc0307_bl = new PETC0307I_BL();
            DataTable dt = new DataTable();
            dt = petc0307_bl.RakutenShop_Combo();

            cboRakutenStore.DisplayMember = "Char1";
            cboRakutenStore.ValueMember = "Key";
            cboRakutenStore.DataSource = dt;
        }

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

        private void F12()
        {
            if (ErrorCheck())
            {
                ShopId = cboRakutenStore.SelectedValue.ToString();
                startdate = txtStartDate.Text.Replace("/", "-");
                enddate = txtEndDate.Text.Replace("/", "-");

                DataTable dtRakuten = new DataTable();
                petc0307_bl = new PETC0307I_BL();
                dtRakuten = petc0307_bl.PECT0301I_Select(ShopId, startdate, enddate);
                if (dtRakuten.Rows.Count > 0)
                {
                    petc0307_bl = new PETC0307I_BL();
                    DataTable dtFolder = new DataTable();
                    dtFolder = petc0307_bl.GetFolder(ShopId);
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
                                csvwriter.WriteCsv(dtRakuten, savedialog.FileName, Encoding.GetEncoding(932));
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
