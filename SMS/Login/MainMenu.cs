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
using SMS_Entity;
using SMS.Inventory;
using SMS.Order;
using System.Deployment.Application;
using Microsoft.VisualBasic.FileIO;

namespace SMS.Login
{
    public partial class frmMainMenu : Form
    {
        public Login_Info loginInfo { get; set; }
        Menu_BL menubl = new Menu_BL();
        MOPE_Entity mope_data;
        string btnName = string.Empty;
        CommonFunctions cf = new CommonFunctions();
        Button btnleftcurrent;
        Button btnrightcurrent;
        
        private void SetDesignerFunction()
        {
            this.KeyPreview = true;
            btnGym1.Click += panelLeft_Click;
            btnGym2.Click += panelLeft_Click;
            btnGym3.Click += panelLeft_Click;
            btnGym4.Click += panelLeft_Click;
            btnGym5.Click += panelLeft_Click;
            btnGym6.Click += panelLeft_Click;
            btnGym7.Click += panelLeft_Click;
            btnGym8.Click += panelLeft_Click;
            btnGym9.Click += panelLeft_Click;
            btnGym10.Click += panelLeft_Click;
            btnGym11.Click += panelLeft_Click;
            btnGym12.Click += panelLeft_Click;
            btnGym13.Click += panelLeft_Click;
            btnGym14.Click += panelLeft_Click;
            btnGym15.Click += panelLeft_Click;
            btnGym16.Click += panelLeft_Click;
            btnGym17.Click += panelLeft_Click;
            btnGym18.Click += panelLeft_Click;
            btnGym19.Click += panelLeft_Click;
            btnGym20.Click += panelLeft_Click;

            btnGym1.MouseEnter += panelLeft_MouseEnter;
            btnGym2.MouseEnter += panelLeft_MouseEnter;
            btnGym3.MouseEnter += panelLeft_MouseEnter;
            btnGym4.MouseEnter += panelLeft_MouseEnter;
            btnGym5.MouseEnter += panelLeft_MouseEnter;
            btnGym6.MouseEnter += panelLeft_MouseEnter;
            btnGym7.MouseEnter += panelLeft_MouseEnter;
            btnGym8.MouseEnter += panelLeft_MouseEnter;
            btnGym9.MouseEnter += panelLeft_MouseEnter;
            btnGym10.MouseEnter += panelLeft_MouseEnter;
            btnGym11.MouseEnter += panelLeft_MouseEnter;
            btnGym12.MouseEnter += panelLeft_MouseEnter;
            btnGym13.MouseEnter += panelLeft_MouseEnter;
            btnGym14.MouseEnter += panelLeft_MouseEnter;
            btnGym15.MouseEnter += panelLeft_MouseEnter;
            btnGym16.MouseEnter += panelLeft_MouseEnter;
            btnGym17.MouseEnter += panelLeft_MouseEnter;
            btnGym18.MouseEnter += panelLeft_MouseEnter;
            btnGym19.MouseEnter += panelLeft_MouseEnter;
            btnGym20.MouseEnter += panelLeft_MouseEnter;

            btnGym1.MouseLeave += panelLeft_MouseLeave;
            btnGym2.MouseLeave += panelLeft_MouseLeave;
            btnGym3.MouseLeave += panelLeft_MouseLeave;
            btnGym4.MouseLeave += panelLeft_MouseLeave;
            btnGym5.MouseLeave += panelLeft_MouseLeave;
            btnGym6.MouseLeave += panelLeft_MouseLeave;
            btnGym7.MouseLeave += panelLeft_MouseLeave;
            btnGym8.MouseLeave += panelLeft_MouseLeave;
            btnGym9.MouseLeave += panelLeft_MouseLeave;
            btnGym10.MouseLeave += panelLeft_MouseLeave;
            btnGym11.MouseLeave += panelLeft_MouseLeave;
            btnGym12.MouseLeave += panelLeft_MouseLeave;
            btnGym13.MouseLeave += panelLeft_MouseLeave;
            btnGym14.MouseLeave += panelLeft_MouseLeave;
            btnGym15.MouseLeave += panelLeft_MouseLeave;
            btnGym16.MouseLeave += panelLeft_MouseLeave;
            btnGym17.MouseLeave += panelLeft_MouseLeave;
            btnGym18.MouseLeave += panelLeft_MouseLeave;
            btnGym19.MouseLeave += panelLeft_MouseLeave;
            btnGym20.MouseLeave += panelLeft_MouseLeave;

            btnPro1.MouseEnter += panelRight_MouseEnter;
            btnPro2.MouseEnter += panelRight_MouseEnter;
            btnPro3.MouseEnter += panelRight_MouseEnter;
            btnPro4.MouseEnter += panelRight_MouseEnter;
            btnPro5.MouseEnter += panelRight_MouseEnter;
            btnPro6.MouseEnter += panelRight_MouseEnter;
            btnPro7.MouseEnter += panelRight_MouseEnter;
            btnPro8.MouseEnter += panelRight_MouseEnter;
            btnPro9.MouseEnter += panelRight_MouseEnter;
            btnPro10.MouseEnter += panelRight_MouseEnter;
            btnPro11.MouseEnter += panelRight_MouseEnter;
            btnPro12.MouseEnter += panelRight_MouseEnter;
            btnPro13.MouseEnter += panelRight_MouseEnter;
            btnPro14.MouseEnter += panelRight_MouseEnter;
            btnPro15.MouseEnter += panelRight_MouseEnter;
            btnPro16.MouseEnter += panelRight_MouseEnter;
            btnPro17.MouseEnter += panelRight_MouseEnter;
            btnPro18.MouseEnter += panelRight_MouseEnter;
            btnPro19.MouseEnter += panelRight_MouseEnter;
            btnPro20.MouseEnter += panelRight_MouseEnter;
            btnPro21.MouseEnter += panelRight_MouseEnter;
            btnPro22.MouseEnter += panelRight_MouseEnter;
            btnPro23.MouseEnter += panelRight_MouseEnter;
            btnPro24.MouseEnter += panelRight_MouseEnter;
            btnPro25.MouseEnter += panelRight_MouseEnter;
            btnPro26.MouseEnter += panelRight_MouseEnter;
            btnPro27.MouseEnter += panelRight_MouseEnter;
            btnPro28.MouseEnter += panelRight_MouseEnter;
            btnPro29.MouseEnter += panelRight_MouseEnter;
            btnPro30.MouseEnter += panelRight_MouseEnter;

            btnPro1.MouseLeave += panelRight_MouseLeave;
            btnPro2.MouseLeave += panelRight_MouseLeave;
            btnPro3.MouseLeave += panelRight_MouseLeave;
            btnPro4.MouseLeave += panelRight_MouseLeave;
            btnPro5.MouseLeave += panelRight_MouseLeave;
            btnPro6.MouseLeave += panelRight_MouseLeave;
            btnPro7.MouseLeave += panelRight_MouseLeave;
            btnPro8.MouseLeave += panelRight_MouseLeave;
            btnPro9.MouseLeave += panelRight_MouseLeave;
            btnPro10.MouseLeave += panelRight_MouseLeave;
            btnPro11.MouseLeave += panelRight_MouseLeave;
            btnPro12.MouseLeave += panelRight_MouseLeave;
            btnPro13.MouseLeave += panelRight_MouseLeave;
            btnPro14.MouseLeave += panelRight_MouseLeave;
            btnPro15.MouseLeave += panelRight_MouseLeave;
            btnPro16.MouseLeave += panelRight_MouseLeave;
            btnPro17.MouseLeave += panelRight_MouseLeave;
            btnPro18.MouseLeave += panelRight_MouseLeave;
            btnPro19.MouseLeave += panelRight_MouseLeave;
            btnPro20.MouseLeave += panelRight_MouseLeave;
            btnPro21.MouseLeave += panelRight_MouseLeave;
            btnPro22.MouseLeave += panelRight_MouseLeave;
            btnPro23.MouseLeave += panelRight_MouseLeave;
            btnPro24.MouseLeave += panelRight_MouseLeave;
            btnPro25.MouseLeave += panelRight_MouseLeave;
            btnPro26.MouseLeave += panelRight_MouseLeave;
            btnPro27.MouseLeave += panelRight_MouseLeave;
            btnPro28.MouseLeave += panelRight_MouseLeave;
            btnPro29.MouseLeave += panelRight_MouseLeave;
            btnPro30.MouseLeave += panelRight_MouseLeave;

            btnPro1.Click += panelRight_Click;
            btnPro2.Click += panelRight_Click;
            btnPro3.Click += panelRight_Click;
            btnPro4.Click += panelRight_Click;
            btnPro5.Click += panelRight_Click;
            btnPro6.Click += panelRight_Click;
            btnPro7.Click += panelRight_Click;
            btnPro8.Click += panelRight_Click;
            btnPro9.Click += panelRight_Click;
            btnPro10.Click += panelRight_Click;
            btnPro11.Click += panelRight_Click;
            btnPro12.Click += panelRight_Click;
            btnPro13.Click += panelRight_Click;
            btnPro14.Click += panelRight_Click;
            btnPro15.Click += panelRight_Click;
            btnPro16.Click += panelRight_Click;
            btnPro17.Click += panelRight_Click;
            btnPro18.Click += panelRight_Click;
            btnPro19.Click += panelRight_Click;
            btnPro20.Click += panelRight_Click;
            btnPro21.Click += panelRight_Click;
            btnPro22.Click += panelRight_Click;
            btnPro23.Click += panelRight_Click;
            btnPro24.Click += panelRight_Click;
            btnPro25.Click += panelRight_Click;
            btnPro26.Click += panelRight_Click;
            btnPro27.Click += panelRight_Click;
            btnPro28.Click += panelRight_Click;
            btnPro29.Click += panelRight_Click;
            btnPro30.Click += panelRight_Click;

            btnF1.Click += btnF1_Click;
            btnF12.Click += btnF12_Click;

            this.FormClosing += frmMainMenu_FormClosing;
            this.KeyDown += frmMainMenu_KeyDown;
        }

        private void panelLeft_MouseEnter(object sender, EventArgs e)
        {
            // when we hoover the button we get this
            (sender as Button).BackColor = Color.LightGreen;
        }

        private void panelLeft_MouseLeave(object sender, EventArgs e)
        {
            // when we again leave the button we get back original color
            if ((sender) as Button != btnleftcurrent)
                (sender as Button).BackColor = Color.FromArgb(192, 255, 192);
        }

        private void panelRight_MouseEnter(object sender, EventArgs e)
        {
            // when we hoover the button we get this
            (sender as Button).BackColor = Color.Khaki;
        }

        private void panelRight_MouseLeave(object sender, EventArgs e)
        {
            // when we again leave the button we get back original color
            if ((sender)as Button != btnrightcurrent)
                (sender as Button).BackColor = Color.FromArgb(255, 224, 192);
        }

        
        public frmMainMenu(MOPE_Entity mope_data, Login_Info loginInfo)
        {
            InitializeComponent();
            if (ApplicationDeployment.IsNetworkDeployed)
                lblVersion.Text = "Version - " + ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString(4);
            SetDesignerFunction();

            this.mope_data = mope_data;
            this.loginInfo = loginInfo;
            Login_BL loginbl = new Login_BL();

            DataTable dt = menubl.Menu_Select(loginInfo);
            if (dt.Rows.Count > 0)
            {
                //lblSCompany.Text = loginInfo.CompanyName = dt.Rows[0][0].ToString();
                //lblSDepartment.Text = loginInfo.DepName = dt.Rows[0][1].ToString();
                //lblSShop.Text = loginInfo.StoreName = dt.Rows[0][2].ToString();
                lblSOperator.Text = loginInfo.OperatorName = dt.Rows[0][3].ToString();
            }

            lblSProcessingDate.Text = DateTime.Today.ToString("yyyy/MM/dd");
            if (mope_data.dtMope.Rows.Count > 0)
            {
                for (int i = 0; i < mope_data.dtMope.Rows.Count; i++)
                {
                    mope_data.GYMNO = mope_data.dtMope.Rows[i][0].ToString();
                    mope_data.PRONO = mope_data.dtMope.Rows[i][3].ToString();
                    mope_data.GYMNM = mope_data.dtMope.Rows[i][2].ToString();
                    //mope_data.GYMCD = mope_data.dtMope.Rows[i]["GYMCD"].ToString();
                    mope_data.PRONM = mope_data.dtMope.Rows[i][5].ToString();
                    mope_data.PROID = mope_data.dtMope.Rows[i]["PROID"].ToString();

                    ButtonText(panelLeft, mope_data, 1);
                }
            }
            else
            {
                mope_data.GYMNO = mope_data.PRONO = mope_data.GYMNM = mope_data.PRONM = string.Empty;
            }
            btnGym1.Focus();
        }

        protected void ButtonText(Panel p, MOPE_Entity MOPEData, int Gym)
        {
            var c = cf.GetAllControls(p);
            for (int i = 0; i < c.Count(); i++)
            {
                Control ctrl = c.ElementAt(i) as Control;
                if (ctrl is Button)
                {
                    if (Gym == 1 && MOPEData.GYMNO != string.Empty)
                    {
                        if (((Button)ctrl).Name == "btnGym" + Convert.ToInt32(MOPEData.GYMNO))
                        {
                            ((Button)ctrl).Text = MOPEData.GYMNM.ToString();
                            ((Button)ctrl).Enabled = true;
                        }
                    }
                    else if (mope_data.PRONO != string.Empty && btnName == "btnGym" + mope_data.GYMNO)
                    {
                        if (((Button)ctrl).Name == "btnPro" + Convert.ToInt32(mope_data.PRONO))
                        {
                            ((Button)ctrl).Text = mope_data.PRONM.ToString();
                            ((Button)ctrl).Enabled = true;
                            ((Button)ctrl).Name = mope_data.PROID.ToString();
                            tt1.SetToolTip(ctrl, mope_data.PROID);
                        }
                    }
                }
            }
        }

        protected void ButtonClick(Button btn)
        {
            btnName = btn.Name;
            if (mope_data.dtMope.Rows.Count > 0)
            {
                for (int i = 0; i < mope_data.dtMope.Rows.Count; i++)
                {
                    mope_data.GYMNO = mope_data.dtMope.Rows[i][0].ToString();
                    mope_data.PRONO = mope_data.dtMope.Rows[i][3].ToString();
                    mope_data.GYMNM = mope_data.dtMope.Rows[i][2].ToString();
                    //mope_data.GYMCD = mope_data.dtMope.Rows[i]["GYMCD"].ToString();
                    mope_data.PRONM = mope_data.dtMope.Rows[i][5].ToString();
                    mope_data.PROID = mope_data.dtMope.Rows[i]["PROID"].ToString();

                    //ButtonText(panelLeft, mope_data, 1);
                    ButtonText(PanelRight, mope_data, 0);
                }
            }
            else
            {
                mope_data.GYMNO = mope_data.PRONO = mope_data.GYMNM = mope_data.PRONM = string.Empty;
            }

            string name = btn.Name;
            string text = btn.Text;
            loginInfo.FormName = btn.Text;
            loginInfo.ItemCode = loginInfo.ItemName = loginInfo.Authority = string.Empty;
            loginInfo.StartMode = "新規";
            Form frm;
            switch (btn.Name)
            {
                case "PSKS0101M":
                    frm = GetForm("PSKS0101M", loginInfo);
                    if (frm == null)
                        new frmPSKS0101M(this.loginInfo).Show();
                    else
                    {
                        frm.Show();
                        frm.WindowState = FormWindowState.Normal;
                    }
                    break;
                case "PSKS0102M":
                    frm = GetForm("PSKS0102M", loginInfo);
                    if (frm == null)
                        new frmPSKS0102M(this.loginInfo).Show();
                    else
                    {
                        frm.Show();
                        frm.WindowState = FormWindowState.Normal;
                    }
                    break;
                case "PSKS0103I":
                    frm = GetForm("PSKS0103I", loginInfo);
                    if (frm == null)
                        new frmPSKS0103I(this.loginInfo).Show();
                    else
                    {
                        frm.Show();
                        frm.WindowState = FormWindowState.Normal;
                    }
                    break;
                case "PSKS0108I":
                    frm = GetForm("PSKS0108I", loginInfo);
                    if (frm == null)
                        new frmPSKS0108I(this.loginInfo).Show();
                    else
                    {
                        frm.Show();
                        frm.WindowState = FormWindowState.Normal;
                    }
                    break;
                case "PSKS0104I":
                    frm = GetForm("PSKS0104I", loginInfo);
                    if (frm == null)
                        new frmPSKS0104I(this.loginInfo).Show();
                    else
                    {
                        frm.Show();
                        frm.WindowState = FormWindowState.Normal;
                    }
                    break;
                case "PSKS0105C":
                    frm = GetForm("PSKS0105C", loginInfo);
                    if (frm == null)
                        new frmPSKS0105C(this.loginInfo).Show();
                    else
                    {
                        frm.Show();
                        frm.WindowState = FormWindowState.Normal;
                    }
                    break;
                case "PSKS0111C":
                    frm = GetForm("PSKS0111Cs", loginInfo);
                    if (frm == null)
                        new frmPSKS0111C(this.loginInfo).Show();
                    else
                    {
                        frm.Show();
                        frm.WindowState = FormWindowState.Normal;
                    }
                    break;
                case "PSKS0110S":
                    frm = GetForm("PSKS0110S", loginInfo);
                    if (frm == null)
                        new frmPSKS0110S(this.loginInfo).Show();
                    else
                    {
                        frm.Show();
                        frm.WindowState = FormWindowState.Normal;
                    }
                    break;
                case "PSKS0114C_1":
                    frm = GetForm("PSKS0114C_1", loginInfo);
                    if (frm == null)
                        new frmPSKS0114C_1(this.loginInfo).Show();
                    else
                    {
                        frm.Show();
                        frm.WindowState = FormWindowState.Normal;
                    }
                    break;
                case "PSKS0119M":
                    frm = GetForm("PSKS0119M", loginInfo);
                    if (frm == null)
                        new frmPSKS0119M(this.loginInfo).Show();
                    else
                    {
                        frm.Show();
                        frm.WindowState = FormWindowState.Normal;
                    }
                    break;
                case "PETC0301I":
                    frm = GetForm("PETC0301I", loginInfo);
                    if (frm == null)
                        new frmPETC0301I(this.loginInfo).Show();
                    else
                    {
                        frm.Show();
                        frm.WindowState = FormWindowState.Normal;
                    }
                    break;

                case "PETC0304I":
                    frm = GetForm("PETC0304I", loginInfo);
                    if (frm == null)
                        new frmPETC0304I(this.loginInfo).Show();
                    else
                    {
                        frm.Show();
                        frm.WindowState = FormWindowState.Normal;
                    }
                    break;
                case "PETC0302I":
                    frm = GetForm("PETC0302I", loginInfo);
                    if (frm == null)
                        new frmPETC0302I(this.loginInfo).Show();
                    else
                    {
                        frm.Show();
                        frm.WindowState = FormWindowState.Normal;
                    }
                    break;
                case "PETC0305I":
                    frm = GetForm("PETC0305I", loginInfo);
                    if (frm == null)
                        new frmPETC0305I(this.loginInfo).Show();
                    else
                    {
                        frm.Show();
                        frm.WindowState = FormWindowState.Normal;
                    }
                    break;
                case "PETC0306I":
                    frm = GetForm("PETC0306I", loginInfo);
                    if (frm == null)
                        new frmPETC0306I(this.loginInfo).Show();
                    else
                    {
                        frm.Show();
                        frm.WindowState = FormWindowState.Normal;
                    }
                    break;
                case "PSKS0120S":
                    frm = GetForm("PSKS0120S", loginInfo);
                    if (frm == null)
                        new frmPSKS0120S(this.loginInfo).Show();
                    else
                    {
                        frm.Show();
                        frm.WindowState = FormWindowState.Normal;
                    }
                    break;
                case "F0001":
                    frm = GetForm("F0001", loginInfo);
                    if (frm == null)
                        new frmF0001(this.loginInfo).Show();
                    else
                    {
                        frm.Show();
                        frm.WindowState = FormWindowState.Normal;
                    }
                    break;
                case "F0002":
                    frm = GetForm("F0002", loginInfo);
                    if (frm == null)
                        new frmF0002(this.loginInfo).Show();
                    else
                    {
                        frm.Show();
                        frm.WindowState = FormWindowState.Normal;
                    }
                    break;
                case "F0003":
                    frm = GetForm("F0003", loginInfo);
                    if (frm == null)
                        new F0003(this.loginInfo).Show();
                    else
                    {
                        frm.Show();
                        frm.WindowState = FormWindowState.Normal;
                    }
                    break;
                case "F0004":
                    frm = GetForm("F0004", loginInfo);
                    if (frm == null)
                        new F0004(this.loginInfo).Show();
                    else
                    {
                        frm.Show();
                        frm.WindowState = FormWindowState.Normal;
                    }
                    break;
                case "F0005":
                    frm = GetForm("F0005", loginInfo);
                    if (frm == null)
                        new frmF0005(this.loginInfo).Show();
                    else
                    {
                        frm.Show();
                        frm.WindowState = FormWindowState.Normal;
                    }
                    break;
                case "PETC0303I":
                    frm = GetForm("PETC0303I", loginInfo);
                    if (frm == null)
                        new frmPETC0303I(this.loginInfo).Show();
                    else
                    {
                        frm.Show();
                        frm.WindowState = FormWindowState.Normal;
                    }
                    break;
                case "PETC0307I":
                    frm = GetForm("PETC0307I", loginInfo);
                    if (frm == null)
                        new frmPETC0307I(this.loginInfo).Show();
                    else
                    {
                        frm.Show();
                        frm.WindowState = FormWindowState.Normal;
                    }
                    break;
                case "PETC0308I":
                    frm = GetForm("PETC0308I", loginInfo);
                    if (frm == null)
                        new frmPETC0308I(this.loginInfo).Show();
                    else
                    {
                        frm.Show();
                        frm.WindowState = FormWindowState.Normal;
                    }
                    break;
                case "PETC0309I":
                    frm = GetForm("PETC0309I", loginInfo);
                    if (frm == null)
                        new frmPETC0309I(this.loginInfo).Show();
                    else
                    {
                        frm.Show();
                        frm.WindowState = FormWindowState.Normal;
                    }
                    break;
                case "PETC0310I":
                    frm = GetForm("PETC0310I", loginInfo);
                    if (frm == null)
                        new frm0310I(this.loginInfo).Show();
                    else
                    {
                        frm.Show();
                        frm.WindowState = FormWindowState.Normal;
                    }
                    break;
                //case "F0006":
                //    frm = GetForm("F0006", loginInfo);
                //    if (frm == null)
                //        new F0006(this.loginInfo).Show();
                //    else
                //    {
                //        frm.Show();
                //        frm.WindowState = FormWindowState.Normal;
                //    }
                //    break;
                case "F0007":
                    frm = GetForm("F0007", loginInfo);
                    if (frm == null)
                        new F0007(this.loginInfo).Show();
                    else
                    {
                        frm.Show();
                        frm.WindowState = FormWindowState.Normal;
                    }
                    break;
                case "F0008":
                    frm = GetForm("F0008", loginInfo);
                    if (frm == null)
                        new F0008(this.loginInfo).Show();
                    else
                    {
                        frm.Show();
                        frm.WindowState = FormWindowState.Normal;
                    }
                    break;
                case "F0029":
                    frm = GetForm("F0029", loginInfo);
                    if (frm == null)
                        new F0029(this.loginInfo).Show();
                    else
                    {
                        frm.Show();
                        frm.WindowState = FormWindowState.Normal;
                    }
                    break;
                case "F0010":
                    frm = GetForm("F0010", loginInfo);
                    if (frm == null)
                        new F0010(this.loginInfo).Show();
                    else
                    {
                        frm.Show();
                        frm.WindowState = FormWindowState.Normal;
                    }
                    break;
                case "F0011":
                    frm = GetForm("F0011", loginInfo);
                    if (frm == null)
                        new F0011(this.loginInfo).Show();
                    else
                    {
                        frm.Show();
                        frm.WindowState = FormWindowState.Normal;
                    }
                    break;
                case "F0012":
                    frm = GetForm("F0012", loginInfo);
                    if (frm == null)
                        new F0012(this.loginInfo).Show();
                    else
                    {
                        frm.Show();
                        frm.WindowState = FormWindowState.Normal;
                    }
                    break;
                case "F0013":
                    frm = GetForm("F0013", loginInfo);
                    if (frm == null)
                        new F0013(this.loginInfo).Show();
                    else
                    {
                        frm.Show();
                        frm.WindowState = FormWindowState.Normal;
                    }
                    break;
                case "F0015":
                    frm = GetForm("F0015", loginInfo);
                    if (frm == null)
                        new F0015(this.loginInfo).Show();
                    else
                    {
                        frm.Show();
                        frm.WindowState = FormWindowState.Normal;
                    }
                    break;
                case "F0016":
                    frm = GetForm("F0016", loginInfo);
                    if (frm == null)
                        new F0016(this.loginInfo).Show();
                    else
                    {
                        frm.Show();
                        frm.WindowState = FormWindowState.Normal;
                    }
                    break;
                case "F0017":
                    frm = GetForm("F0017", loginInfo);
                    if (frm == null)
                        new F0017(this.loginInfo).Show();
                    else
                    {
                        frm.Show();
                        frm.WindowState = FormWindowState.Normal;
                    }
                    break;
                case "F0019":
                    frm = GetForm("F0019", loginInfo);
                    if (frm == null)
                        new F0019(this.loginInfo).Show();
                    else
                    {
                        frm.Show();
                        frm.WindowState = FormWindowState.Normal;
                    }
                    break;
                case "F0020":
                    frm = GetForm("F0020", loginInfo);
                    if (frm == null)
                        new F0020(this.loginInfo).Show();
                    else
                    {
                        frm.Show();
                        frm.WindowState = FormWindowState.Normal;
                    }
                    break;
                case "F0021":
                    frm = GetForm("F0021", loginInfo);
                    if (frm == null)
                        new frmF0021(this.loginInfo).Show();
                    else
                    {
                        frm.Show();
                        frm.WindowState = FormWindowState.Normal;
                    }
                    break;
                case "F0022":
                    frm = GetForm("F0021", loginInfo);
                    if (frm == null)
                        new F0022(this.loginInfo).Show();
                    else
                    {
                        frm.Show();
                        frm.WindowState = FormWindowState.Normal;
                    }
                    break;
                case "F0023":
                    frm = GetForm("F0021", loginInfo);
                    if (frm == null)
                        new F0023(this.loginInfo).Show();
                    else
                    {
                        frm.Show();
                        frm.WindowState = FormWindowState.Normal;
                    }
                    break;
                case "F0024":
                    frm = GetForm("F0024", loginInfo);
                    if (frm == null)
                        new F0024(this.loginInfo).Show();
                    else
                    {
                        frm.Show();
                        frm.WindowState = FormWindowState.Normal;
                    }
                    break;
                case "F0026":
                    frm = GetForm("F0026", loginInfo);
                    if (frm == null)
                        new F0026(this.loginInfo).Show();
                    else
                    {
                        frm.Show();
                        frm.WindowState = FormWindowState.Normal;
                    }
                    break;
                case "F0027":
                    frm = GetForm("F0027", loginInfo);
                    if (frm == null)
                        new F0027(this.loginInfo).Show();
                    else
                    {
                        frm.Show();
                        frm.WindowState = FormWindowState.Normal;
                    }
                    break;
                case "F0028":
                    frm = GetForm("F0028", loginInfo);
                    if (frm == null)
                        new frmF0028(this.loginInfo).Show();
                    else
                    {
                        frm.Show();
                        frm.WindowState = FormWindowState.Normal;
                    }
                    break;
                case "F0030":
                    frm = GetForm("F0030", loginInfo);
                    if (frm == null)
                        new F0030(this.loginInfo).Show();
                    else
                    {
                        frm.Show();
                        frm.WindowState = FormWindowState.Normal;
                    }
                    break;
                case "F0031":
                    frm = GetForm("F0031", loginInfo);
                    if (frm == null)
                        new F0031(this.loginInfo).Show();
                    else
                    {
                        frm.Show();
                        frm.WindowState = FormWindowState.Normal;
                    }
                    break;
            }
        }

        public Form GetForm(string formID, Login_Info loginInfo)
        {
            FormCollection frms = Application.OpenForms;
            foreach (Form frm in frms)
            {
                if (frm.Text.Equals(formID))
                {
                    return frm;
                }
            }

            return null;
        }

        private void functionKeyPress(byte index)
        {
            switch (index)
            {
                case 1:
                    if (cf.DSP_MSG("Q003", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty).Equals("1"))
                        Environment.Exit(1);
                    break;
                case 12:
                    Application.Restart();
                    break;
            }
        }

        private void frmMainMenu_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F1 || e.KeyCode == Keys.F12)
                functionKeyPress(Convert.ToByte(e.KeyCode.ToString().Substring(1)));
        }

        private void btnF1_Click(object sender, EventArgs e)
        {
            functionKeyPress(1);
        }

        private void btnF12_Click(object sender, EventArgs e)
        {
            functionKeyPress(12);
        }

        private void panelLeft_Click(object sender, EventArgs e)
        {
            var c = cf.GetAllControls(PanelRight);
            int j = 0;
            for (int i = 0; i < c.Count(); i++)
            {
                Control ctrl = c.ElementAt(i) as Control;
                if (ctrl is Button)
                {
                    j++;
                    ((Button)ctrl).Text = "";
                    ((Button)ctrl).Enabled = false;
                    ((Button)ctrl).Name = "btnPro" + (c.Count() - (j));
                }
            }

            c = cf.GetAllControls(panelLeft);
            j = 0;
            for (int i = 0; i < c.Count(); i++)
            {
                Control ctrl = c.ElementAt(i) as Control;
                if (ctrl is Button)
                {
                    ((Button)ctrl).BackColor = Color.FromArgb(192, 255, 192);
                }
            }

            Button btn = sender as Button;
            
            ButtonClick(btn);
            btn.BackColor = Color.LightGreen;
            btnleftcurrent = btn;
        }

        private void panelRight_Click(object sender, EventArgs e)
        {
            var c = cf.GetAllControls(PanelRight);
            int j = 0;
            for (int i = 0; i < c.Count(); i++)
            {
                Control ctrl = c.ElementAt(i) as Control;
                if (ctrl is Button)
                {
                    ((Button)ctrl).BackColor = Color.FromArgb(255, 224, 192);
                }
            }

            Button btn = sender as Button;
            ButtonClick(btn);
            btn.BackColor = Color.Khaki;
            btnrightcurrent = btn;
        }

        private void frmMainMenu_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                if (cf.DSP_MSG("Q003", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty).Equals("1"))
                    e.Cancel = false;
                else
                    e.Cancel = true;
            }
        }

        public DataTable CSVToTable(string filePath,string firstColHeader)
        {
            DataTable csvData = new DataTable();
            int count = 1;
            try
            {
                using (TextFieldParser csvReader = new TextFieldParser(filePath, Encoding.GetEncoding(932), true))
                {
                    csvReader.SetDelimiters(new string[] { "," });
                    csvReader.HasFieldsEnclosedInQuotes = true;
                    //read column names
                    string[] colFields = csvReader.ReadFields();
                    char c = 'A';

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
                            DataColumn datacolumn = new DataColumn(c++.ToString());
                            datacolumn.AllowDBNull = true;
                            csvData.Columns.Add(datacolumn);
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
            catch (Exception)
            {
            }
            return csvData;
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            DataTable dt = CSVToTable(@"E:\Import\globeride.csv", "2");
        }
    }
}
