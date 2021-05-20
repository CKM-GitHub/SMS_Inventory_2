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

namespace SMS.Login
{
    public partial class frmLoginPassword : Form
    {
        CommonFunctions com = new CommonFunctions();
        MOPE_Entity mope_data = new MOPE_Entity();
        Login_BL loginbl = new Login_BL();

        public frmLoginPassword(MOPE_Entity mopeData)
        {
            InitializeComponent();
            txtTANCD.Text = mopeData.TANCD;
            txtTANNM.Text = mopeData.TANNM;
            txtHideTextBox.Text = mopeData.PASWD;
            SetDesignerFunction();
        }

        private void SetDesignerFunction()
        {
            txtTANCD.Enabled = false;
            txtTANNM.Enabled = false;
            btnF12.Enabled = false;
            this.StartPosition = FormStartPosition.CenterScreen;

            this.Load += frmLoginPassword_Load;
            this.KeyDown += frmKeyDown;
            txtOldPASWD.KeyDown += txtOldPASWD_KeyDown;
            txtNewPASWD.KeyDown += txtNewPASWD_KeyDown;
            txtConfirmPASWD.KeyDown += txtConfirmPASWD_KeyDown;
            btnF1.Click += btnF1_Click;
            btnF12.Click += btnF12_Click;
            txtHideTextBox.Visible = false;
        }

        private void frmLoginPassword_Load(object sender, EventArgs e)
        {
            this.KeyPreview = true;
            this.KeyDown += frmKeyDown;
        }

        public void frmKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SendKeys.Send("{TAB}");
            }
        }

        public bool TextboxRequireCheck(TextBox txt, string str)
        {
            bool b = false;
            if (string.IsNullOrWhiteSpace(txt.Text.ToString()))
            {
                com.DSP_MSG("E017", str, string.Empty, string.Empty, string.Empty, string.Empty).Equals("1");
                txt.Focus();
                b = true;
            }
            return b;
        }

        private void txtOldPASWD_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && !TextboxRequireCheck(txtOldPASWD, "これまでのパスワード"))
            {
                if (txtHideTextBox.Text != txtOldPASWD.Text.ToString())
                {
                    com.DSP_MSG("E001", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty).Equals("1");
                }
            }
        }

        private void txtNewPASWD_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                TextboxRequireCheck(txtOldPASWD, "新しいパスワード");
            }
        }

        private void txtConfirmPASWD_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (txtConfirmPASWD.Text != string.Empty && (!TextboxRequireCheck(txtConfirmPASWD, "確認　再入力")))
                {
                    if (txtNewPASWD.Text.ToString() == txtConfirmPASWD.Text.ToString())
                    {
                        btnF12.Enabled = true;
                        mope_data.PASWD = txtConfirmPASWD.Text.ToString();
                        btnF12.Focus();
                    }
                    else
                    {
                        com.DSP_MSG("E028", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty).Equals("1");
                    }
                }
            }
        }

        private void btnF1_Click(object sender, EventArgs e)
        {
            functionKeypress(1);
        }

        private void btnF12_Click(object sender, EventArgs e)
        {
            functionKeypress(12);
        }

        private void functionKeypress(byte index)
        {
            switch (index)
            {
                case 1:
                    if (com.DSP_MSG("C014", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty).Equals("1"))
                        this.Close();
                    break;
                case 12:
                    mope_data.TANCD = txtTANCD.Text.ToString();
                    bool b = loginbl.MOPE_UPDATE(mope_data);
                    if (b)
                    {
                        txtNewPASWD.Text = txtConfirmPASWD.Text = string.Empty;
                        btnF12.Enabled = false;
                        txtHideTextBox.Text = mope_data.PASWD;
                        com.DSP_MSG("I002", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty).Equals("1");
                    }
                    break;
            }
        }
    }
}
