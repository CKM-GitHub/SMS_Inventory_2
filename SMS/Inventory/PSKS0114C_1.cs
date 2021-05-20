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

namespace SMS.Inventory
{
    public partial class frmPSKS0114C_1 : SMS_Base
    {
        Login_Info loginInfo;
        M_MultiPorpose_Entity mme;
        PSKS0114C_BL psks0114cbl;
        L_Log_Entity lle;

        #region constructor
        public frmPSKS0114C_1()
        {
            InitializeComponent();
        }

        public frmPSKS0114C_1(Login_Info loginInfo) 
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
        private void frmPSKS0114C_1_Load(object sender, EventArgs e)
        {
            psks0114cbl = new PSKS0114C_BL();
            StatusControl();
            SetDesignFunction();
        }

        /// <summary>
        /// Control Flag at Form Load
        /// </summary>
        private void StatusControl()
        {
            mme = new M_MultiPorpose_Entity();
            mme = psks0114cbl.M_MultiPorpose_Num_Select();
            if (mme.Num1 == "1")
            {
                lblSKSMode.Text = "処理実行中";
                lblSKSMode.BackColor = Color.FromArgb(0, 176, 240);
            }
            else
            {
                lblSKSMode.Text = "処理停止中";
                lblSKSMode.BackColor = Color.Yellow;
            }
        }

        /// <summary>
        /// Design Setting
        /// </summary>
        private void SetDesignFunction()
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
            FunctionButtonDisabled(12);
            FormName = "在庫情報反映";
            lblMode.Visible = false;
        }

        /// <summary>
        /// Control F1 Function Click 
        /// </summary>
        /// <param name="btnIndex"></param>
        public override void functionClick(int btnIndex)
        {
            switch (btnIndex)
            {
                case 1:
                    break;
            }
        }

        /// <summary>
        /// Button Click For FLag Change
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnChange_Click(object sender, EventArgs e)
        {
            var btn = (Button)sender;
            mme = new M_MultiPorpose_Entity();
            lle = new L_Log_Entity();

            mme.ID = "108";
            //mme.Num1 = "1";
            mme.InsertOperator = lle.InsertOperator = loginInfo.OperatorCode;
            mme.InsertDateTime = lle.OperateDate = DateTime.Now.ToString();
            mme.Program =lle.Program = "PSKS0114C";
            lle.PC = loginInfo.PcName;

            //lle.OperateMode = "処理開始";
           
            if (btn.Text == "開始")
            {
                if (lblSKSMode.Text != "処理実行中")
                {
                    lblSKSMode.Text = "処理実行中";
                    lblSKSMode.BackColor = Color.FromArgb(0, 176, 240);
                    mme.Num1 = "1";
                    psks0114cbl.M_MultiPorposeNum1_Update(mme, lle);
                }
            }
            else
            {
                if (lblSKSMode.Text != "処理停止中")
                {
                    lblSKSMode.Text = "処理停止中";
                    lblSKSMode.BackColor = Color.Yellow;
                    mme.Num1 = "0";
                    psks0114cbl.M_MultiPorposeNum1_Update(mme, lle);
                }
            }
        }
    }
}
