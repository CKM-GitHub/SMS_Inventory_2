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
using Sms_Prog;

namespace SMS.Inventory
{
    public partial class frmPSKS0105C : SMS_Base
    {
        Login_Info loginInfo;
        PSKS0105C_BL psks0105cbl;
        M_MultiPorpose_Entity mme;
        L_Log_Entity lle;

        #region constructor
        /// <summary>
        /// Default
        /// </summary>
        public frmPSKS0105C()
        {
            InitializeComponent();
        }

        public frmPSKS0105C(Login_Info loginInfo)
            : base(loginInfo)
        {
            InitializeComponent();
            this.loginInfo = loginInfo;
        }
        #endregion

        /// <summary>
        /// form load
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void frmPSKS0105C_Load(object sender, EventArgs e)
        {
            psks0105cbl = new PSKS0105C_BL();
            Control();

            FunctionButtonDisabled(2);
            FunctionButtonDisabled(3);
            FunctionButtonDisabled(4);
            FunctionButtonDisabled(5);
            FunctionButtonDisabled(6);
            FunctionButtonDisabled(7);
            FunctionButtonDisabled(8);
            FunctionButtonDisabled(9);
            FunctionButtonDisabled(10);
            FunctionButtonDisabled(12);
            lblMode.Visible = false;
            FormName = "販売・入出庫情報取得";

            dgvPSKS0105C.DisabledColumn("colNo,colInsertDateTime,colHanbaiShohinCD,colVM_HANBAI_SHOHIN,colM3,colM1,colSu,colM2,colTekiyou");
            dgvPSKS0105C.DisabledSorting();
        }


        /// <summary>
        /// Control Button color (Blue and Yellow)
        /// </summary>
        private void Control()
        {
            DataTable dtNum1 = new DataTable();

            lblNum1.Text = psks0105cbl.M_MultiPorpose_Num3_Select() + "日間の履歴を保持しています";

            DataTable dt = new DataTable();
            dt = psks0105cbl.M_MultiPorpose_MultiSelect();
            if (dt.Rows.Count > 0)
            {
                DataTable dtWeb = dt.Select("ID=102").CopyToDataTable();
                if (dtWeb.Rows.Count > 0)
                {
                    if (dtWeb.Rows[0]["Num1"].ToString() == "1")
                    {
                        lblWebMode.Text = "処理実行中";
                        lblWebMode.BackColor = Color.FromArgb(0, 176, 240);
                    }
                    else
                    {
                        lblWebMode.Text = "処理停止中";
                        lblWebMode.BackColor = Color.Yellow;
                    }

                }

                DataTable dtInventory = dt.Select("ID=103").CopyToDataTable();
                if (dtInventory.Rows.Count > 0)
                {
                    if (dtInventory.Rows[0]["Num1"].ToString() == "1")
                    {
                        lblInventoryMode.Text = "処理実行中";
                        lblInventoryMode.BackColor = Color.FromArgb(0, 176, 240);
                    }
                    else
                    {
                        lblInventoryMode.Text = "処理停止中";
                        lblInventoryMode.BackColor = Color.Yellow;
                    }
                }

                DataTable dtStore = dt.Select("ID=107").CopyToDataTable();
                if (dtStore.Rows.Count > 0)
                {
                    if (dtStore.Rows[0]["Num1"].ToString() == "1")
                    {
                        lblStoreMode.Text = "処理実行中";
                        lblStoreMode.BackColor = Color.FromArgb(0, 176, 240);
                    }
                    else
                    {
                        lblStoreMode.Text = "処理停止中";
                        lblStoreMode.BackColor = Color.Yellow;
                    }
                }
            }
        }


        /// <summary>
        /// Control to run or not PSKS0106B console
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnWeb_Click(object sender, EventArgs e)
        {
            var btnWeb = (Button)sender;
            mme = new M_MultiPorpose_Entity();
            lle = new L_Log_Entity();

            mme.ID = "102";
            mme.InsertOperator = lle.InsertOperator = loginInfo.OperatorCode;
            mme.InsertDateTime = lle.OperateDate = DateTime.Now.ToString();
            mme.Program = "PSKS0105C";
            lle.PC = loginInfo.PcName;

            if (btnWeb.Text == "開始")
            {
                if (lblWebMode.Text != "処理実行中")
                {
                    lblWebMode.Text = "処理実行中";
                    lblWebMode.BackColor = Color.FromArgb(0, 176, 240);
                    mme.Num1 = "1";
                    psks0105cbl.M_MultiPorpose_Num1_Update(mme, lle);
                }
            }
            else
            {
                if (lblWebMode.Text != "処理停止中")
                {
                    lblWebMode.Text = "処理停止中";
                    lblWebMode.BackColor = Color.Yellow;
                    mme.Num1 = "0";
                    psks0105cbl.M_MultiPorpose_Num1_Update(mme, lle);
                }
            }
        }

        /// <summary>
        /// Control to run or not PSKS0107B console
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnInventory_Click(object sender, EventArgs e)
        {
            var btnInventory = (Button)sender;
            mme = new M_MultiPorpose_Entity();
            lle = new L_Log_Entity();

            mme.ID = "103";
            mme.InsertOperator = lle.InsertOperator = loginInfo.OperatorCode;
            mme.InsertDateTime = lle.OperateDate = DateTime.Now.ToString();
            lle.Program = "PSKS0105C";
            lle.PC = loginInfo.PcName;
            if (btnInventory.Text == "開始")
            {
                if (lblInventoryMode.Text != "処理実行中")
                {
                    lblInventoryMode.Text = "処理実行中";
                    lblInventoryMode.BackColor = Color.FromArgb(0, 176, 240);
                    mme.Num1 = "1";
                    psks0105cbl.M_MultiPorpose_Num1_Update(mme, lle);
                }
            }
            else
            {
                if (lblInventoryMode.Text != "処理停止中")
                {
                    lblInventoryMode.Text = "処理停止中";
                    lblInventoryMode.BackColor = Color.Yellow;
                    mme.Num1 = "0";
                    psks0105cbl.M_MultiPorpose_Num1_Update(mme, lle);
                }
            }
        }


        /// <summary>
        /// Control to run or not PSKS0113B console
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnStore_Click(object sender, EventArgs e)
        {
            var btnStore = (Button)sender; 
            mme = new M_MultiPorpose_Entity();
            lle = new L_Log_Entity();

            mme.ID = "107";
            mme.InsertOperator = lle.InsertOperator = loginInfo.OperatorCode;
            mme.InsertDateTime = lle.OperateDate = DateTime.Now.ToString();
            lle.Program = "PSKS0105C";
            lle.PC = loginInfo.PcName;

            if (btnStore.Text == "開始")
            {
                if (lblStoreMode.Text != "処理実行中")
                {
                    lblStoreMode.Text = "処理実行中";
                    lblStoreMode.BackColor = Color.FromArgb(0, 176, 240);
                    mme.Num1 = "1";
                    psks0105cbl.M_MultiPorpose_Num1_Update(mme, lle);
                }
            }
            else
            {
                if (lblStoreMode.Text != "処理停止中")
                {
                    lblStoreMode.Text = "処理停止中";
                    lblStoreMode.BackColor = Color.Yellow;
                    mme.Num1 = "0";
                    psks0105cbl.M_MultiPorpose_Num1_Update(mme, lle);
                }
            }
        }

        /// <summary>
        /// handle f1 to f12 click event
        /// implement base virtual function
        /// </summary>
        /// <param name="btnIndex"></param>
        public override void functionClick(int btnIndex)
        {
            switch (btnIndex)
            {
                case 1:
                    break;
                case 11:
                    Search();
                    break;
               
            }
        }

        /// <summary>
        /// Searching
        /// Repair In 02182019 by PTK
        /// </summary>
        private  void Search() 
        {

            Cursor.Current = Cursors.WaitCursor;

            DataTable dt = new DataTable();
            dt = psks0105cbl.T_Ukebarai_Select();
            dgvPSKS0105C.DataSource = dt;
            Cursor.Current = Cursors.Default;
            btnShow.Focus();
        }

        //public async Task<bool> GridView()
        //{
        //    Func<bool> function = new Func<bool>(() => F11());
        //    return await Task.Factory.StartNew<bool>(function);
        //}

        //private bool F11()
        //{
        //    if (btnShow.InvokeRequired)
        //    {
        //        btnShow.Invoke(new MethodInvoker(delegate
        //        {
        //            DataTable dt = new DataTable();
        //            dt = psks0105cbl.T_Ukebarai_Select();
        //            dgvPSKS0105C.DataSource = dt;
        //        }));
        //    }
        //    return true;
        //}

        /// <summary>
        /// Search
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnShow_Click(object sender, EventArgs e)
        {
            Search();
        }

       
    }
}
