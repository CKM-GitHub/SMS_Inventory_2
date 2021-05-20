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
using System.Threading;
using Sms_Prog;
using Newtonsoft.Json;
using SMS_DL;
using System.IO;
using System.Net;

namespace SMS.Inventory
{
    public partial class frmPSKS0111C : SMS_Base
    {
        Login_Info loginInfo;
        PSKS0111C_BL psks0111cbl;
        M_MultiPorpose_Entity mme;
        L_Log_Entity lle;
        MOPE_Entity mope_data;
       public static Base_DL bdl =new Base_DL();
        string RenkeiDateTime = string.Empty;

        #region constructor
        /// <summary>
        /// Default
        /// </summary>
        public frmPSKS0111C()
        {
            InitializeComponent();
        }

        public frmPSKS0111C(Login_Info loginInfo)
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
        private void frmPSKS0111C_Load(object sender, EventArgs e)
        {
            psks0111cbl = new PSKS0111C_BL();
            Control();
            GridView();
            SetDesignFunction();
            dgvTSZRK.DisabledColumn("colNo,colRenkeiDateTime,colCOUNT");
            dgvTSZRK_Detail.DisabledColumn("*");
            dgvTSZRK.DisabledSorting();
            dgvTSZRK_Detail.DisabledSorting();
            //if (this.loginInfo.OperatorCode != "0001")
            //    btnSKSRecovery.Visible = false;
        }

        private void SetDesignFunction()
        {
            FormName = "在庫SKS連携画面";
            lblMode.Visible = false;
            FunctionButtonDisabled(2);
            FunctionButtonDisabled(3);
            FunctionButtonDisabled(4);
            FunctionButtonDisabled(6);
            FunctionButtonDisabled(7);
            FunctionButtonDisabled(8);
            FunctionButtonDisabled(9);
            FunctionButtonDisabled(10);
            FunctionButtonDisabled(12);
            this.KeyDown += btnUpdate_KeyDown;
        }


        /// <summary>
        /// Display inventory count by Date
        /// </summary>
        private void GridView()
        {
            DataTable dt_TSZRK = new DataTable();
            dt_TSZRK = psks0111cbl.T_SKSZaikoRireki_Select();
            dgvTSZRK.DataSource = dt_TSZRK;
        }

        #region Searching
        /// <summary>
        /// Searching
        /// 02182019 repaired by PTk
        /// </summary>
        private  void F11()  
        {
            if (!string.IsNullOrWhiteSpace(RenkeiDateTime))
            {
                Cursor.Current = Cursors.WaitCursor;

                DataTable dtTSZRK_Detail = new DataTable();
                dtTSZRK_Detail = psks0111cbl.T_SKSZaikoRireki_DetailSelect(RenkeiDateTime);
                dgvTSZRK_Detail.DataSource = dtTSZRK_Detail;

                Cursor.Current = Cursors.Default;
                btnShowGrid.Focus();
            }
            else
                dgvTSZRK_Detail.DataSource = null;
        }

        /// <summary>
        /// Display for Gridview
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnShowGrid_Click(object sender, EventArgs e)
        {
            F11();
        }

        #endregion

        #region Control to run or not console
        private void Control()
        {
            DataTable dt_Num = new DataTable();
            mme = psks0111cbl.M_MultiPorpose_Num_Select();
            lblNum3.Text = mme.Num3 + "日間の履歴を保持しています";

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
        /// Change button color
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSKS(object sender, EventArgs e)
        {
            var btnSKS = (Button)sender;
            mme = new M_MultiPorpose_Entity();
            lle = new L_Log_Entity();

            mme.ID = "109";
            mme.InsertOperator = lle.InsertOperator = loginInfo.OperatorCode;
            mme.InsertDateTime = lle.OperateDate = DateTime.Now.ToString();
            lle.PC = loginInfo.PcName;
            mme.Program = "PSKS0111C";

            if (btnSKS.Text == "開始") 
            {
                if (lblSKSMode.Text != "処理実行中")
                {
                    lblSKSMode.Text = "処理実行中";
                    lblSKSMode.BackColor = Color.FromArgb(0, 176, 240);
                    mme.Num1 = "1";
                    psks0111cbl.M_MultiPorpose_Num1_Update(mme, lle);
                    
                }

            }
            
            else
            {
                if (lblSKSMode.Text != "処理停止中")
                {
                    lblSKSMode.Text = "処理停止中";
                    lblSKSMode.BackColor = Color.Yellow;
                    mme.Num1 = "0";
                    psks0111cbl.M_MultiPorpose_Num1_Update(mme, lle);
                }
            }

        }

        #endregion

        /// <summary>
        /// handle f1 to f12 click event
        /// implement base virtual function
        /// </summary>
        /// <param name="btnIndex"></param>
        public override void functionClick(int btnIndex)
        {
            switch (btnIndex)
            {
                case 5:
                    Clear();
                    break;
                case 1:
                    break;
                case 11:
                    F11();
                    break;
                case 12:
                    F12();
                    break;
                
            }
        }

        /// <summary>
        /// Clear grid
        /// </summary>
        private void Clear()
        {
            GridView();
            dgvTSZRK_Detail.DataSource = null;         
        }

        #region Refresh form
        private void btnUpdate_Click(object sender, EventArgs e)
        {
            F12();
        }

        public  void F12()
        {
            psks0111cbl = new PSKS0111C_BL();
            Control();
            GridView();
            SetDesignFunction();
            dgvTSZRK.DisabledColumn("colNo,colRenkeiDateTime,colCOUNT");
            dgvTSZRK_Detail.DisabledColumn("*");
            dgvTSZRK.DisabledSorting();
            dgvTSZRK_Detail.DisabledSorting();
            RenkeiDateTime = string.Empty;
        }

        #endregion

        #region keydown event
        public void btnUpdate_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F12)
            {
                F12();
            }
        }

        #endregion

        #region Cell click event of Gridview
        private void dgvTSZRK_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            Maintained_CellClick(sender,e);
        }

        private void dgvTSZRK_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

            Maintained_CellClick(sender, e);
        }

        protected void Maintained_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if ((sender as DataGridView).CurrentCell is DataGridViewCheckBoxCell)
            {
                if ((Convert.ToBoolean(dgvTSZRK.Rows[e.RowIndex].Cells[e.ColumnIndex].EditedFormattedValue) == true))
                {
                    foreach (DataGridViewRow row1 in dgvTSZRK.Rows)
                    {
                        DataGridViewCheckBoxCell chk1 = row1.Cells[e.ColumnIndex] as DataGridViewCheckBoxCell;
                        chk1.Value = chk1.FalseValue;
                        if (row1.Index % 2 == 0)
                        {
                            row1.DefaultCellStyle.BackColor = Color.White;
                        }
                        else
                        {
                            row1.DefaultCellStyle.BackColor = Color.FromArgb(221, 235, 247);
                        }
                    }

                    dgvTSZRK.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = true;
                    RenkeiDateTime = dgvTSZRK.Rows[e.RowIndex].Cells[2].Value.ToString().Replace("/", "-");
                    dgvTSZRK.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Yellow;
                }
                else
                {
                    RenkeiDateTime = string.Empty;
                    dgvTSZRK.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.White;
                    dgvTSZRK.ClearSelection();
                }

                dgvTSZRK.EndEdit();
            }
        }

        #endregion

        #region Resend inventory data to SKS 
        private void btnSKSRecovery_Click(object sender, EventArgs e)
        {
            string optCode = this.loginInfo.OperatorCode;
            if (optCode == "0001")
            {
                if (!string.IsNullOrWhiteSpace(RenkeiDateTime))
                {
                    DataTable dtTSZRK_Detail = new DataTable();
                    dtTSZRK_Detail=psks0111cbl.ResendInventoryToSKS(RenkeiDateTime);
                    string Js = datatableToJason(dtTSZRK_Detail);
                    PostJS(Js);
                }
            }
        }

        /// <summary>
        /// Change to Json format
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public string datatableToJason(DataTable dt)
        {
            string JS = string.Empty;
            JS = "{\"inventorynew\":" + JsonConvert.SerializeObject(dt) + "}";
            return JS;
        }

        /// <summary>
        /// Send to SKS
        /// </summary>
        /// <param name="re"></param>
        public static void PostJS(string re)
        {
            var baseAddress = bdl.path_108i("Sks_Path", "http://www.capital-sys.net/RCM_Capital/WebForms/Import/Json_Import.ashx?User=system&Password=capital");
            var http = (HttpWebRequest)WebRequest.Create(new Uri(baseAddress));
            http.Accept = "application/json";
            http.ContentType = "application/json";
            http.Method = "POST";
            string parsedContent = re;
            Encoding encoding = Encoding.UTF8;
            //ASCIIEncoding encoding = new ASCIIEncoding();
            Byte[] bytes = encoding.GetBytes(parsedContent);
            http.Timeout = System.Threading.Timeout.Infinite;

            http.ReadWriteTimeout = System.Threading.Timeout.Infinite;
            Stream newStream = http.GetRequestStream();
            newStream.Write(bytes, 0, bytes.Length);
            newStream.Close();

            var response = http.GetResponse();
            var stream = response.GetResponseStream();
            var sr = new StreamReader(stream);
            var content = sr.ReadToEnd();
        }

        #endregion
    }
}
