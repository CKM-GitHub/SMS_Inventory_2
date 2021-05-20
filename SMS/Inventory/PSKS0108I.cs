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
using System.Threading;
using System.IO;
using SMS_DL;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace SMS.Inventory
{
    public partial class frmPSKS0108I : SMS_Base
    {
        public Base_DL bd = new Base_DL();
        Login_Info loginInfo;
        PSKS0108I_BL psks0108ibl;
        T_SugorakuInport_Entity tsie;
        M_MultiPorpose_Entity mme;
        public frmPSKS0108I()
        {
            InitializeComponent();
        }
        public frmPSKS0108I(Login_Info loginInfo)
            : base(loginInfo)
        {
            InitializeComponent();
            this.loginInfo = loginInfo;
            this.FormName = "在庫情報取得";
            SetDesignerFunction();
            lblMode.Visible = false;
        }

        private void SetDesignerFunction()
        {
            FunctionButtonDisabled(2);
            FunctionButtonDisabled(3);
            FunctionButtonDisabled(4);
            FunctionButtonDisabled(6);
            FunctionButtonDisabled(7);
            FunctionButtonDisabled(8);
            FunctionButtonDisabled(9);
            FunctionButtonDisabled(10);
            FunctionButtonDisabled(11);
            FunctionButtonDisabled(12);
        }

        private void frmPSKS0108I_Load(object sender, EventArgs e)
        {
            psks0108ibl = new PSKS0108I_BL();
            buttonEnable_Check();
            
        }
        private void ActivateTimer()
        {
         
                bool IsTicked = false;
                string flg;
                var t = new System.Windows.Forms.Timer();
                t.Interval = 3600000; // it will Tick in 1 hour- seconds
                t.Tick += (s, e) =>
                {
                    if (!CheckDate()) // I hour 
                    {
                        WindowState = FormWindowState.Maximized;
                        BringToFront();
                        this.TopMost = true;
                        Focus();
                        // this.ShowOver();
                        // this.BringToFront();
                        flg = DSP_MSG("I301", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                        if (flg == "1")
                            IsTicked = true;
                        else
                            IsTicked = false;
                        t.Stop();
                    }
                };
                t.Start();
           // }
        }
        private void ShowOver()
        {
            try
            {
                Process[] localByName = Process.GetProcessesByName("販売管理");
                if (localByName.Count() > 0)
                {
                    this.Show();
                    // MessageBox.Show("Got Count");
                    IntPtr handle = localByName[0].MainWindowHandle;
                    ShowWindow(handle, SW_SHOWMAXIMIZED);
                    SetForegroundWindow(handle);
                    return;
                }
            }
            catch { }
        }
        private const int SW_SHOWMAXIMIZED = 3;
        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport("user32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern bool SetForegroundWindow(IntPtr hwnd);
        private void buttonEnable_Check()
        {
            DataTable dt = psks0108ibl.T_SugorakuInport_SELECT();
            if (dt.Rows.Count > 0)
            {
                DateTime startDate = DateTime.Parse(dt.Rows[0]["InportStartTime"].ToString());
                DateTime EndDate = DateTime.Parse(dt.Rows[0]["InportEndTime"].ToString());

                if (!DateTime.Now.ToString("yyyy/MM/dd").Equals(Convert.ToDateTime(startDate).ToString("yyyy/MM/dd")))
                {
                    btnStage3.Enabled = false;
                    btnStage4.Enabled = false;
                }
                else
                {
                    btnStage3.Enabled = true;
                    btnStage4.Enabled = true;

                    if (startDate > EndDate)
                    {
                        btnStage4.Enabled = false;
                    }
                    //else if (DateTime.Now.ToString("yyyy/MM/dd").Equals(Convert.ToDateTime(startDate).ToString("yyyy/MM/dd")))
                    else
                    {
                        btnStage4.Enabled = true;
                    }
                }
            }
        }

        private void btnWeb_Click(object sender, EventArgs e)
        {
            DateTime datetime = System.DateTime.Now;
            tsie = new T_SugorakuInport_Entity();
            tsie.InportStartTime = datetime.ToString();
            tsie.InportStartKBN = "1";

            mme = new M_MultiPorpose_Entity();
            mme.Num1 = "0";
            mme.InsertOperator = loginInfo.OperatorCode;
            mme.InsertDateTime = datetime.ToString();

            if (psks0108ibl.btnWeb_Update(tsie, mme))
            {
                lblStage2.BackColor = Color.Yellow;
                txtStartTime.Text = datetime.ToString("HH:mm");
            }

            //***Bug51 Need Button Enable/Disable on Step1 click,not only WebSaveClick
            buttonEnable_Check();
            ActivateTimer();
        }

        private void btnWebSave_Click(object sender, EventArgs e)
        {
            if (DateTimeNullCheck(txtStartTime))
            {
                tsie = new T_SugorakuInport_Entity();
                tsie.InportStartTime = System.DateTime.Now.Date.ToShortDateString() + " " + txtStartTime.Text.ToString();
                tsie.InportStartKBN = "2";
                psks0108ibl.T_SugorakuInport_Update(tsie);
                lblStage2.BackColor = Color.Yellow;
                buttonEnable_Check();
            }
        }

        private void btnStage3_Click(object sender, EventArgs e)
        {
            DateTime datetime = DateTime.Now;
            tsie = new T_SugorakuInport_Entity();
            tsie.InportEndTime = datetime.ToString();
            tsie.InportEndKBN = "1";
            psks0108ibl = new PSKS0108I_BL();

            if (psks0108ibl.T_SugorakuInport_Update(tsie))
            {
                lblStage2.BackColor = Color.FromArgb(255, 242, 204);
                txtEndTime.Text = datetime.ToString("HH:mm");
            }

            //***Bug51 Need Button Enable/Disable on Step3 click,not only btnStage3Save_Click
            buttonEnable_Check();
        }

        private void btnStage3Save_Click(object sender, EventArgs e)
        {
            if (DateTimeNullCheck(txtEndTime))
            {
                tsie = new T_SugorakuInport_Entity();
                //***BugNo48 System Error Occurs becasue of parse txtEndTime to db,correct --> txtEndTime.Text
                tsie.InportEndTime = System.DateTime.Now.Date.ToShortDateString() + " " + txtEndTime.Text.ToString();
                tsie.InportEndKBN = "2";
                psks0108ibl.T_SugorakuInport_Update(tsie);
                lblStage2.BackColor = Color.FromArgb(255, 242, 204);
                buttonEnable_Check();
            }
        }

        //***BugNo57 *** All button Disable After Click Step4 button
        private void Step1()
        {
            lblStage5.BackColor = Color.Yellow;
            lblStage7.BackColor = Color.FromArgb(255, 242, 204);
            if (btnStage4.InvokeRequired)
            {
                btnStage4.Invoke(new MethodInvoker(delegate
                {
                    btnWeb.Enabled = false;
                    btnStage3.Enabled = false;
                    btnStage4.Enabled = false;
                    btnWebSave.Enabled = false;
                    btnStage3Save.Enabled = false;
                }));
            }
        }

        private void Step2()
        {
            lblStage5.BackColor = Color.FromArgb(255, 242, 204);
            lblStage6.BackColor = Color.Yellow;
        }

        private bool UpdateMultiPurpose()
        {
            Thread.Sleep(5 * 1000);//to stop stage6 with yellow color for 5 sec
            string dateTime = System.DateTime.Now.ToString();
            mme = new M_MultiPorpose_Entity();
            mme.Num1 = mme.Key = "1";
            mme.InsertOperator = loginInfo.OperatorCode;
            mme.InsertDateTime = dateTime;

            mme.ID = "108";
            if (psks0108ibl.M_MultiPorpose_Num1_Update(mme))
            {
                mme.ID = "109";
                if (psks0108ibl.M_MultiPorpose_Num1_Update(mme))
                {

                    lblStage6.BackColor = Color.FromArgb(255, 242, 204);
                    lblStage7.BackColor = Color.Yellow;

                    return true;
                }
                return false;
            }
            return false;
        }

        public async Task<bool> Insert()
        {
            Func<bool> function = new Func<bool>(() => InsertData());
            return await Task.Factory.StartNew<bool>(function);
        }

        public async Task<bool> M_Update()
        {
            Func<bool> function = new Func<bool>(() => UpdateMultiPurpose());
            return await Task.Factory.StartNew<bool>(function);
        }

        private bool CheckDate()
        {
            DataTable dt = psks0108ibl.T_SugorakuInport_SELECT();
            if (dt.Rows.Count > 0)
            {
                DateTime startDate = DateTime.Parse(dt.Rows[0]["InportStartTime"].ToString());
                DateTime EndDate = DateTime.Parse(dt.Rows[0]["InportEndTime"].ToString());
                if (startDate > EndDate)
                    return false;
            }
                 
            return true;
        }
        private async void btnStage4_Click(object sender, EventArgs e)
        {
            if (!CheckDate())
            {
                DSP_MSG("I301", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                return;
            }
            Task task = new Task(new Action(Step1));
            task.Start();
            bool result = await Insert();
            if (result)
            {
                task = new Task(new Action(Step2));
                task.Start();

                bool result1 = await M_Update();

                if (result1)
                {
                    buttonEnable_Check();
                    btnWeb.Enabled = true;
                    btnWebSave.Enabled = true;
                    btnStage3Save.Enabled = true;
                }
            }
        }
        //end

        private bool InsertData()
        {
            DateTime datetime = DateTime.Now;
            psks0108ibl = new PSKS0108I_BL();
            L_Log_Entity LogData = new L_Log_Entity();
            LogData.OperateDate = datetime.ToString("yyyy/MM/dd");
            LogData.OperateTime = datetime.ToString("HH:mm");
            LogData.InsertOperator = loginInfo.OperatorCode;
            LogData.Program = "PSKS0108I";
            LogData.PC = loginInfo.PcName;

            #region copy file

            //String sourcePath = bd.path_108i("S_path", @"\\192.168.0.14\ablesoft\CONPHAS-GENERAL\DATA");
            //String targetPath = bd.path_108i("T_Path", @"\\192.168.0.5\Dropbox (Capital Group)\Capital backup\systemdata\inventory\ShopData\108I");
            // string fileName = "FreeDAT.mdb";
            //string s = System.IO.Path.Combine(sourcePath, fileName);
            //string destFile = System.IO.Path.Combine(targetPath, fileName);

            string shopFlag = "0";
           
            //if(File.Exists(s))
            //{
            //    System.IO.File.Copy(s, destFile, true);
            //    shopFlag = "1";
            //}

            #endregion

            bool result = psks0108ibl.T_RONRI_ZAIKO_Insert(LogData, datetime.ToString(),shopFlag);
            //bool result = true;
            //if (File.Exists(destFile))
            //    File.Delete(destFile);

            return result;

        }//更新処理１

        public override void functionClick(int index)
        {
            switch (index)
            {
                case 1:
                    break;
                case 5:
                    Clear();
                    break;
            }
        }
        private void Clear()
        {
            txtStartTime.Text = string.Empty;
            txtEndTime.Text = string.Empty;
        }
    }
}
