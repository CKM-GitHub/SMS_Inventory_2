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

namespace SMS.PopUp
{
    public partial class frmZaikoHenkou : Form
    {
        Login_Info loginInfo;
        public bool flgExist = false;
        public string oldJisha = string.Empty;
        public string oldToyonaka = string.Empty;
        //T_Zaiko_Entity tze;
        L_Log_Entity lle;
        CommonFunctions cf;
        ZaikoHenkan_BL zhbl;

        public T_Zaiko_Entity tze { get; set; }

        #region Constructor
        public frmZaikoHenkou()
        {
            InitializeComponent();
        }

        public frmZaikoHenkou(T_Zaiko_Entity ZaikoData, Login_Info loginInfo)
        {
            InitializeComponent();
            this.loginInfo = loginInfo;
            txtSKUCD.Text = ZaikoData.HanbaiShohinCD;
            txtJANCD.Text = ZaikoData.JanCD;
            txtItemName.Text = ZaikoData.ItemName;
            txtColor.Text = ZaikoData.Color;
            txtSize.Text = ZaikoData.Size;
            txtCapitalZaiko.Text = ZaikoData.CapitalZaiko;
            txtToyonakaZaiko.Text = ZaikoData.ToyonakaZaiko;
            txtIshibashiZaiko.Text = ZaikoData.IshibashiZaiko;
            txtEsakaZaiko.Text = ZaikoData.EsakaZaiko;
            txtSannomiyaZaiko.Text = ZaikoData.SannomiyaZaiko;

        }
        #endregion

        /// <summary>
        /// Form Load
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmZaikoHenkou_Load(object sender, EventArgs e)
        {
            zhbl = new ZaikoHenkan_BL();
            cf = new CommonFunctions();
        }

        /// <summary>
        /// KeyDown Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmZaikoHenkou_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F1)
            {
                this.Close();
            }
            else if (e.KeyCode == Keys.F12)
            {
                F12();
            }
        }

        #region Button Click
        private void btnF1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnF12_Click(object sender, EventArgs e)
        {
            F12();
        }
        #endregion

        /// <summary>
        /// InsertUpdate Data
        /// </summary>
        private void F12()
        {
            Dictionary<Control, string> dic = new Dictionary<Control, string>();
            dic.Add(txtCapitalZaiko, string.Empty);
            dic.Add(txtToyonakaZaiko, string.Empty);
            dic.Add(txtIshibashiZaiko, string.Empty);
            dic.Add(txtEsakaZaiko, string.Empty);
            dic.Add(txtSannomiyaZaiko, string.Empty);
            if (ErrorCheck(dic))
            {
                if (cf.DSP_MSG("Q101", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty).Equals("1"))
                {
                    oldJisha = txtCapitalZaiko.Text;
                    oldToyonaka = txtToyonakaZaiko.Text;
                    InsertUpdateData();
                    this.Close();
                }
                else
                {
                    flgExist = true;
                    this.Close();
                }
            }
        }

        private void InsertUpdateData()
        {
            string DateAndTime = System.DateTime.Now.ToString();
            tze = new T_Zaiko_Entity();
            tze.HanbaiShohinCD = txtSKUCD.Text;
            tze.JanCD = txtJANCD.Text;
            tze.CapitalZaiko = txtCapitalZaiko.Text;
            tze.ToyonakaZaiko = txtToyonakaZaiko.Text;
            tze.IshibashiZaiko = txtIshibashiZaiko.Text;
            tze.EsakaZaiko = txtEsakaZaiko.Text;
            tze.SannomiyaZaiko = txtSannomiyaZaiko.Text;
            tze.UpdateDateTime = DateAndTime;

            lle = new L_Log_Entity();
            lle.OperateDate = DateAndTime;
            lle.InsertOperator = loginInfo.OperatorCode;
            lle.Program = "PSKS110S";
            lle.PC = loginInfo.PcName;
            lle.OperateMode = "在庫変更";
            lle.KeyItem = txtSKUCD.Text;

            if (zhbl.T_Zaiko_INSERT_UPDATE(tze,lle))
            {
                
            }
        }

        /// <summary>
        /// Error Check for F12 Click
        /// </summary>
        /// <param name="dic"></param>
        /// <returns></returns>
        private bool ErrorCheck(Dictionary<Control, string> dic)
        {
            foreach (KeyValuePair<Control, string> kvp in dic)
            {
                if (kvp.Key is TextBox)
                {
                    TextBox txt = kvp.Key as TextBox;
                    if (string.IsNullOrWhiteSpace(txt.Text))
                    {
                        txt.Text = "0";
                    }
                    else if (Convert.ToInt32(txt.Text.Replace (",","")) < 0)
                    {
                        cf.DSP_MSG("E109", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                        txt.Focus();
                        return false;
                    }
                }
            }
            return true;
        }
    }
}
