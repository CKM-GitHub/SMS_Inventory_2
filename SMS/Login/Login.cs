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
using System.IO;
using System.Deployment.Application;
using SMS_DL;

namespace SMS.Login
{
    public partial class frmLogin : Form
    {

        MKAI_Entity mkaiEntity = new MKAI_Entity();
        MOPE_Entity mopeData = new MOPE_Entity();
        Login_BL loginbl = new Login_BL();
        Login_Info logininfo = new Login_Info();
        CommonFunctions com = new CommonFunctions();

        public frmLogin()
        {
            InitializeComponent();

            Base_DL change_dbstring = new Base_DL();
//change_dbstring.Change_DBVersion();
            InstallUpdateSyncWithInfo();
            SetDesignerFunction();
            SetMaxLength();
            //SetTabIndex();
            txtTANCD.Select();

            if (ApplicationDeployment.IsNetworkDeployed)
                lblVersion.Text = "Version - " +  ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString(4);
        }

        public IEnumerable<int> Integers()
{
    yield return 1;
    yield return 2;
    yield return 4;
    yield return 8;
    yield return 16;
    yield return 16777216;
}
        public void Consumer()
        {
            foreach (int i in Integers())
            {
                Console.WriteLine(i.ToString());
            }
        }
        private void Login_Load(object sender, EventArgs e)
        {



           // Consumer();



            bnF12.Enabled = false;
            btnChange.Enabled = false;
            txtKAICD.Focus();
            lblSProcessingDate.Text = DateTime.Today.ToString("yyyy/MM/dd");
        }


        private void SetDesignerFunction()
        {
            this.KeyPreview = true;     
            this.Load += Login_Load;
            //txtKAICD.KeyDown += txtKAICD_KeyDown;
            txtTANCD.KeyDown += txtTANCD_KeyDown;
            txtPASWD.KeyDown += txtPASWD_KeyDown;
            btnChange.Click += btnChange_Click;
            //bnF1.Click += bnF1_Click;
            //bnF12.Click += bnF12_Click;
            lblSProcessingDate.Text = DateTime.Today.ToString("yyyy/MM/dd");
            txtPASWD.PasswordChar = '●';

            bnF12.Enabled = false;
            btnChange.Enabled = false;
            txtTANCD.Focus(); 
        }

        private void InstallUpdateSyncWithInfo()
        {
            Base_DL change_dbstring = new Base_DL();
            UpdateCheckInfo info = null;

            if (ApplicationDeployment.IsNetworkDeployed)
            {
                change_dbstring.Change_DBVersion();
                ApplicationDeployment ad = ApplicationDeployment.CurrentDeployment;

                try
                {
                    info = ad.CheckForDetailedUpdate();

                }
                catch (DeploymentDownloadException dde)
                {
                    MessageBox.Show("The new version of the application cannot be downloaded at this time. \n\nPlease check your network connection, or try again later. Error: " + dde.Message);
                    return;
                }
                catch (InvalidDeploymentException ide)
                {
                    MessageBox.Show("Cannot check for a new version of the application. The ClickOnce deployment is corrupt. Please redeploy the application and try again. Error: " + ide.Message);
                    return;
                }
                catch (InvalidOperationException ioe)
                {
                    MessageBox.Show("This application cannot be updated. It is likely not a ClickOnce application. Error: " + ioe.Message);
                    return;
                }

                if (info.UpdateAvailable)
                {
                    Boolean doUpdate = true;

                    //if (!info.IsUpdateRequired)
                    //{
                    //    DialogResult dr = MessageBox.Show("An update is available. Would you like to update the application now?", "Update Available", MessageBoxButtons.OKCancel);
                    //    if (!(DialogResult.OK == dr))
                    //    {
                    //        doUpdate = false;
                    //    }
                    //}
                    //else
                    //{
                    //    // Display a message that the app MUST reboot. Display the minimum required version.
                    //    MessageBox.Show("This application has detected a mandatory update from your current " +
                    //        "version to version " + info.MinimumRequiredVersion.ToString() +
                    //        ". The application will now install the update and restart.",
                    //        "Update Available", MessageBoxButtons.OK,
                    //        MessageBoxIcon.Information);
                    //}

                    if (doUpdate)
                    {
                        try
                        {
                            ad.Update();

                

                            MessageBox.Show("The application has been upgraded, and will now restart.");
                            Application.Restart();
                        }
                        catch (DeploymentDownloadException dde)
                        {
                           // change_dbstring.path_108i("IsChanged_db", "0");
                            MessageBox.Show("Cannot install the latest version of the application. \n\nPlease check your network connection, or try again later. Error: " + dde);
                            return;
                        }
                    }
                }

            }
            else {
             //   change_dbstring.path_108i("IsChanged_db", "0");
            }
        }

        private void SetTabIndex()
        {
            txtKAICD.TabIndex = 0;
            txtTANCD.TabIndex = 1;
            txtPASWD.TabIndex = 2;
            bnF12.TabIndex = 3;
            btnChange.TabIndex = 4;
            bnF1.TabIndex = 5;
        }

        private void SetMaxLength()
        {
            txtKAICD.MaxLength = 2;
            txtTANCD.MaxLength = 4;
            txtPASWD.MaxLength = 10;
        }

        private void btnChange_Click(object sender, EventArgs e)
        {
            bnF12.Enabled = false;
            bnF1.Enabled = false;
            frmLoginPassword loginPassword = new frmLoginPassword(mopeData);
            loginPassword.ShowDialog();
            bnF1.Enabled = true; bnF12.Enabled = true;
        }

        private void txtKAICD_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && (!TextboxRequireCheck(txtKAICD, "会社CD")))
            {
                mkaiEntity.KAICD = txtKAICD.Text.ToString();
                DataTable dt = loginbl.MKAI_SELECT(mkaiEntity);//Select Company Name
                if (dt.Rows.Count > 0)
                {
                    lblJSYRN.Text = dt.Rows[0][0].ToString();
                    txtTANCD.Focus();
                }
                else
                {
                    com.DSP_MSG("E012", "会社CD", string.Empty, string.Empty, string.Empty, string.Empty).Equals("1");
                    lblJSYRN.Text = string.Empty;
                }
            }
        }

        public bool TextboxRequireCheck(TextBox txt, string str)
        {
            if (string.IsNullOrWhiteSpace(txt.Text.ToString()))
            {
                com.DSP_MSG("E017", str, string.Empty, string.Empty, string.Empty, string.Empty).Equals("1");
                txt.Focus();
                return true;
            }
            bnF1.Enabled = true;
            return false;
        }

        private void txtTANCD_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && !TextboxRequireCheck(txtTANCD, "オペレータCD"))
            {
                txtPASWD.Focus();
                //mopeData.KAICD = txtKAICD.Text.ToString();
                //mopeData.TANCD = txtTANCD.Text.ToString();
                //DataTable dt = loginbl.MOPE_SELECT(mopeData);// Select Operator Name and Password
                //if (dt.Rows.Count > 0)
                //{
                //    mopeData.PASWD = dt.Rows[0][0].ToString();
                //    mopeData.TANNM = dt.Rows[0][6].ToString();

                //    txtPASWD.Focus();
                //    bnF12.Enabled=true;
                //}
                //else
                //{
                //    MessageBox.Show("Invalid Operator Code!");
                //    //com.DSP_MSG("E003", txtTANCD.Text, string.Empty, string.Empty, string.Empty, string.Empty).Equals("1");
                //}
                //bnF1.Enabled = true;
            }
        }

        private void txtPASWD_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && !TextboxRequireCheck(txtPASWD, "パスワード"))
            {
                logininfo.CompanyCode = mopeData.KAICD = txtKAICD.Text.ToString();
                logininfo.OperatorCode = mopeData.TANCD = txtTANCD.Text.ToString();
                DataTable dt = loginbl.MOPE_SELECT(mopeData);
                if (dt.Rows.Count > 0)
                {
                    mopeData.PASWD = dt.Rows[0][0].ToString();
                    mopeData.KNGCD = dt.Rows[0][1].ToString();
                    mopeData.MNUCD = dt.Rows[0][2].ToString();
                    logininfo.StoreCode = mopeData.TMPCD = dt.Rows[0][4].ToString();
                    logininfo.DepCode = mopeData.BMNCD = dt.Rows[0][5].ToString();
                }

                if (mopeData.PASWD != string.Empty && mopeData.PASWD == txtPASWD.Text.ToString())
                {
                    bnF12.Enabled = true;
                    //btnChange.Enabled = true;
                    ////PanelFunction.Focus();
                    functionKeypress(12);

                }
                else
                {
                   // btnF12.Enabled = false;
                    bnF1.Enabled = false;
                    com.DSP_MSG("E001", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty).Equals("1");
                }
            }
            bnF1.Enabled = true;
            //int d = Convert.ToInt32("ds");
        }

        private void functionKeypress(byte index)
        {
            switch (index)
            {
                case 1:
                    if (com.DSP_MSG("Q003", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty).Equals("1"))
                        this.Close();
                    break;
                case 12:
                    if (bnF12.Enabled)
                    {


                        logininfo.CompanyCode = mopeData.KAICD = txtKAICD.Text.ToString();
                        logininfo.OperatorCode = mopeData.TANCD = txtTANCD.Text.ToString();
                        ////Log and store to notepad   in path //     ~/Resources/store_operator    // Added by Phyoe Gyi 03-26-2019     -- to make temp store in note like session for tracing_log table

                        //string path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location.Replace("\\bin\\Debug", "")) + @"\Resources\";
                        //File.WriteAllText(Path.Combine(path, "store_operator"), "");
                        //using (StreamWriter writer = File.AppendText(Path.Combine(path, "store_operator")))
                        //   {
                           
                        //       writer.WriteLine(logininfo.CompanyCode + "|" + logininfo.OperatorCode);
                        //       writer.Close();
                        //   }

                        DataTable dt1 = loginbl.MOPE_SELECT(mopeData);
                        if (dt1.Rows.Count > 0)
                        {
                            mopeData.PASWD = dt1.Rows[0][0].ToString();
                            mopeData.KNGCD = dt1.Rows[0][1].ToString();
                            mopeData.MNUCD = dt1.Rows[0][2].ToString();
                            logininfo.StoreCode = mopeData.TMPCD = dt1.Rows[0][4].ToString();
                            logininfo.DepCode = mopeData.BMNCD = dt1.Rows[0][5].ToString();
                        }
                        if (mopeData.PASWD != string.Empty && mopeData.PASWD == txtPASWD.Text.ToString())
                        {
                            
                            DataTable dt = loginbl.SMNU_SELECT(mopeData);
                            //mopeData = new MOPE_Data();
                            mopeData.dtMope = dt;
                            logininfo.PcLoginName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
                            logininfo.PcName = System.Environment.MachineName;
                            logininfo.CallerID = string.Empty;
                            logininfo.StartMode = string.Empty;
                            logininfo.Authority = string.Empty;
                            logininfo.StartMode = "新規";
                            logininfo.ProcessDate = lblSProcessingDate.Text;
                            logininfo.ProcessClass = "1";
                            bnF12.Enabled = false;
                            btnChange.Enabled = false;
                       

                            
                           
                            this.Hide();
                            frmMainMenu mm = new frmMainMenu(mopeData, logininfo);
                            mm.ShowDialog();
                            this.Close();
                            Environment.Exit(1);
                        }
                        else {
                           

                                com.DSP_MSG("E017", txtPASWD.Text.ToString(), string.Empty, string.Empty, string.Empty, string.Empty).Equals("1");
                                txtPASWD.Focus();

                                bnF1.Enabled = true;
                        }
                        break;
                    }
                    break;
            }
        }

        private void bnF1_Click(object sender, EventArgs e)
        {
            functionKeypress(1);
        }

        private void bnF12_Click(object sender, EventArgs e)
        {
          //  bnF12.Enabled = true;
            functionKeypress(12);
        }

        private void bnF1_Click_2(object sender, EventArgs e)
        {
            functionKeypress(1); 
        }

        private void bnF12_Click_2(object sender, EventArgs e)
        {
          //  bnF12.Enabled = true;
            functionKeypress(12);
        }

        private void frmLogin_KeyDown(object sender, KeyEventArgs e)
        {
            if ((e.KeyCode == Keys.F1 && (bnF1.Cursor == Cursors.Hand)) )
                      
            {
               // MessageBox.Show("Hi this is F1 and F12");
                functionKeypress(1);
                
            }
            else if (e.KeyCode == Keys.F12 && (bnF12.Cursor == Cursors.Hand))
            {
               // bnF12.Enabled = true;
                functionKeypress(12);
            }

        }
    }
}
