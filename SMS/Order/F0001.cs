using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data;
using SMS_Entity;

namespace SMS.Order
{
    public partial class frmF0001 : SMS_Base
    {
        Login_Info loginInfo;

        public frmF0001()
        {
            InitializeComponent();
        }

        public frmF0001(Login_Info loginInfo)
            : base(loginInfo)
        {
            InitializeComponent();
            this.loginInfo = loginInfo;
            SetDesignerFunction();
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

            dgv2.DisabledColumn("colNo");
            FormMode = "登録";
        }

        private void BindCombo()
        {
            #region 店舗
            DataTable dtStore = new DataTable();
            dtStore.Columns.Add("StoreID");
            dtStore.Columns.Add("StoreName");

            dtStore.Rows.Add("-1", "");
            dtStore.Rows.Add("0", "Web店");
            dtStore.Rows.Add("1", "豊中店");
            dtStore.Rows.Add("2", "石橋店");
            dtStore.Rows.Add("3", "三宮店");
            dtStore.Rows.Add("4", "江坂店");

            dtStore.AcceptChanges();
            comboStore.SelectedIndex = -1;

            this.comboStore.DisplayMember = "StoreName";
            this.comboStore.ValueMember = "StoreID";
            this.comboStore.DataSource = dtStore;
            #endregion

            #region 会員区分
            DataTable dtMClass = new DataTable();
            dtMClass.Columns.Add("MClassID");
            dtMClass.Columns.Add("MClassName");

            dtMClass.Rows.Add("-1", "");
            dtMClass.Rows.Add("0", "WEB顧客");
            dtMClass.Rows.Add("1", "店舗顧客");

            dtMClass.AcceptChanges();
            cboMemberClass.SelectedIndex = -1;

            this.cboMemberClass.DisplayMember = "MClassName";
            this.cboMemberClass.ValueMember = "MClassID";
            this.cboMemberClass.DataSource = dtMClass;
            #endregion

            #region DM発行区分
            DataTable dtDMClass = new DataTable();
            dtDMClass.Columns.Add("DMClassID");
            dtDMClass.Columns.Add("DMClassName");

            dtDMClass.Rows.Add("-1", "");
            dtDMClass.Rows.Add("0", "する");
            dtDMClass.Rows.Add("1", "しない");

            dtDMClass.AcceptChanges();
            cboDMClass.SelectedIndex = -1;

            this.cboDMClass.DisplayMember = "DMClassName";
            this.cboDMClass.ValueMember = "DMClassID";
            this.cboDMClass.DataSource = dtDMClass;
            #endregion

            #region 締日
            DataTable dtCDate = new DataTable();
            dtCDate.Columns.Add("CDateID");
            dtCDate.Columns.Add("CDateName");

            dtCDate.Rows.Add("-1", "");
            dtCDate.Rows.Add("0", "10日");
            dtCDate.Rows.Add("1", "15日");
            dtCDate.Rows.Add("2", "20日");
            dtCDate.Rows.Add("3", "月末");

            dtCDate.AcceptChanges();
            cboClosingDate.SelectedIndex = -1;

            this.cboClosingDate.DisplayMember = "CDateName";
            this.cboClosingDate.ValueMember = "CDateID";
            this.cboClosingDate.DataSource = dtCDate;
            #endregion

            #region 請求月
            DataTable dtBMonth = new DataTable();
            dtBMonth.Columns.Add("BMonthID");
            dtBMonth.Columns.Add("BMonthName");

            dtBMonth.Rows.Add("-1", "");
            dtBMonth.Rows.Add("0", "10日");
            dtBMonth.Rows.Add("1", "15日");
            dtBMonth.Rows.Add("2", "20日");
            dtBMonth.Rows.Add("3", "月末");

            dtBMonth.AcceptChanges();
            cboBillingMonth.SelectedIndex = -1;

            this.cboBillingMonth.DisplayMember = "BMonthName";
            this.cboBillingMonth.ValueMember = "BMonthID";
            this.cboBillingMonth.DataSource = dtBMonth;
            #endregion

            #region 請求日
            DataTable dtBDate = new DataTable();
            dtBDate.Columns.Add("BDateID");
            dtBDate.Columns.Add("BDateName");

            dtBDate.Rows.Add("-1", "");
            dtBDate.Rows.Add("0", "10日");
            dtBDate.Rows.Add("1", "15日");
            dtBDate.Rows.Add("2", "20日");
            dtBDate.Rows.Add("3", "月末");

            dtBDate.AcceptChanges();
            cboBillingDate.SelectedIndex = -1;

            this.cboBillingDate.DisplayMember = "BDateName";
            this.cboBillingDate.ValueMember = "BDateID";
            this.cboBillingDate.DataSource = dtBDate;
            #endregion

            #region 入金方法
            DataTable dtPayment = new DataTable();
            dtPayment.Columns.Add("PaymentID");
            dtPayment.Columns.Add("PaymentName");

            dtPayment.Rows.Add("-1", "");
            dtPayment.Rows.Add("0", "現金");
            dtPayment.Rows.Add("1", "カード支払");
            dtPayment.Rows.Add("2", "銀行振込");
            dtPayment.Rows.Add("3", "コンビニ支払");

            dtPayment.AcceptChanges();
            cboPayment.SelectedIndex = -1;

            this.cboPayment.DisplayMember = "PaymentName";
            this.cboPayment.ValueMember = "PaymentID";
            this.cboPayment.DataSource = dtPayment;
            #endregion

            #region 税区分
            DataTable dtTax = new DataTable();
            dtTax.Columns.Add("TaxID");
            dtTax.Columns.Add("TaxName");

            dtTax.Rows.Add("-1", "");
            dtTax.Rows.Add("0", "課税");
            dtTax.Rows.Add("1", "非課税");

            dtTax.AcceptChanges();
            comboBox11.SelectedIndex = -1;

            this.comboBox11.DisplayMember = "TaxName";
            this.comboBox11.ValueMember = "TaxID";
            this.comboBox11.DataSource = dtTax;
            #endregion


            #region 要注意顧客
            DataTable dtPay = new DataTable();
            dtPay.Columns.Add("PayID");
            dtPay.Columns.Add("PayName");

            dtPay.Rows.Add("-1", "");
            dtPay.Rows.Add("0", "過去債権トラブル");
            dtPay.Rows.Add("1", "過去取引要注意");

            dtPay.AcceptChanges();
            cboPayAttention.SelectedIndex = -1;

            this.cboPayAttention.DisplayMember = "PayName";
            this.cboPayAttention.ValueMember = "PayID";
            this.cboPayAttention.DataSource = dtPay;
            #endregion

            #region 要確認顧客
            DataTable dt = new DataTable();
            dt.Columns.Add("ID");
            dt.Columns.Add("Name");

            dt.Rows.Add("-1", "");
            dt.Rows.Add("0", "請求書同封");
            dt.Rows.Add("1", "キャンセル率高");
            dt.Rows.Add("2", "交換率高");

            dt.AcceptChanges();
            comboBox13.SelectedIndex = -1;

            this.comboBox13.DisplayMember = "Name";
            this.comboBox13.ValueMember = "ID";
            this.comboBox13.DataSource = dt;
            #endregion
        }

        public void MallBind()
        {
            DataTable dtmall = new DataTable();
            dtmall.Columns.Add("MallID");
            dtmall.Columns.Add("MallName");

            dtmall.Rows.Add("0", string.Empty);
            dtmall.Rows.Add("1", "Yahoo");
            dtmall.Rows.Add("2", "楽天");
            dtmall.Rows.Add("3", "Amazon");
            dtmall.Rows.Add("4", "Wowma");
            dtmall.AcceptChanges();

            DataGridViewComboBoxColumn col = (DataGridViewComboBoxColumn)dgv1.Columns["colMall"];
            col.DataPropertyName = "MallID";
            col.DisplayMember = "MallName";

            col.ValueMember = "MallID";



            ((DataGridViewComboBoxColumn)dgv1.Columns["colMall"]).DataSource = dtmall;


        }

        public void MallBind1()
        {
            DataTable dtmall = new DataTable();
            dtmall.Columns.Add("MallID");
            dtmall.Columns.Add("MallName");

            dtmall.Rows.Add("0", string.Empty);
            dtmall.Rows.Add("1", "Yahoo");
            dtmall.Rows.Add("2", "楽天");
            dtmall.Rows.Add("3", "Amazon");
            dtmall.Rows.Add("4", "Wowma");
            dtmall.AcceptChanges();

            DataGridViewComboBoxColumn col = (DataGridViewComboBoxColumn)dgv2.Columns["colMall1"];
            col.DataPropertyName = "MallID";
            col.DisplayMember = "MallName";

            col.ValueMember = "MallID";



            ((DataGridViewComboBoxColumn)dgv2.Columns["colMall1"]).DataSource = dtmall;


        }

        public void MallCombo()
        {
            DataTable malldt = new DataTable();

            malldt.Columns.Add("Email", typeof(string));
            //malldt.Columns.Add("ItemProperty", typeof(string));

            malldt.Rows.Add();

            MallBind();
            dgv1.DataSource = malldt;
            //EnableGridColumn(smS_GridView1.Columns["Email"], true);
        }

        private void frmF0001_Load(object sender, EventArgs e)
        {
            MallCombo();
            MallBind1();
            BindCombo();
            //BindGrid1();
            //BindGrid2();
            FormName = "顧客マスター";
            
        }

        private void dgv1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            ComboBox cbo = e.Control as ComboBox;
            if (cbo != null)
            {
                cbo.SelectedIndexChanged -= new EventHandler(ComBoBox_SelectedIndexChange);
                cbo.SelectedIndexChanged += new EventHandler(ComBoBox_SelectedIndexChange);
            }
        }

        private void ComBoBox_SelectedIndexChange(Object sender, EventArgs e)
        {
            //ComboBox cbo = sender as ComboBox;
            //int row = dgv1.CurrentCell.RowIndex;
            //string item = cbo.Text;
            //if (item == "Yahoo")
            //{
            //    dgv1.Rows[row].Cells["colEmail"].Value = "test@co.jp";
            //}
            //else if (item == "楽天")
            //{
            //    dgv1.Rows[row].Cells["colEmail"].Value = "data@co.jp";
            //}
            //else if (item == "Amazon")
            //{
            //    dgv1.Rows[row].Cells["colEmail"].Value = "aaa@co.jp";
            //}
            //else if (item == "Wauma")
            //{
            //    dgv1.Rows[row].Cells["colEmail"].Value = "bbb@co.jp";
            //}
            //else
            //{
            //    dgv1.Rows[row].Cells["colEmail"].Value = "";
            //}
        }

        //private void BindGrid2()
        //{
        //DataTable dtgv2 = new DataTable();
        //dtgv2.Columns.Add("No");
        //dtgv2.Columns.Add("Name");
        //dtgv2.Columns.Add("PhoneNumber");
        //dtgv2.Columns.Add("MailAddress");
        //dtgv2.Columns.Add("Address1");
        //dtgv2.Columns.Add("Address2");
        //dtgv2.Columns.Add("MallID");
        //dtgv2.Columns.Add("RegistrationDate");

        //string[] row1 = new string[] { "1", "ＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸＸ", "09012345678", "xxxxxxxx@xxxxxxxx.co.jp", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "1", "9999/99/99" };
        //string[] row2 = new string[] { "2", "", "09012345678", "xxxxxxxx@xxxxxxxx.co.jp", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "2","9999/99/99" };
        //string[] row3 = new string[] { "3", "", "09012345678", "xxxxxxxx@xxxxxxxx.co.jp", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "4","9999/99/99" };
        //string[] row4 = new string[] { "4", "", "09012345678", "xxxxxxxx@xxxxxxxx.co.jp", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "1","9999/99/99" };
        //string[] row5 = new string[] { "5", "", "09012345678", "xxxxxxxx@xxxxxxxx.co.jp", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "2","9999/99/99" };
        //string[] row6 = new string[] { "6", "", "09012345678", "xxxxxxxx@xxxxxxxx.co.jp", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "4","9999/99/99" };

        //dtgv2.Rows.Add(row1);
        //dtgv2.Rows.Add(row2);
        //dtgv2.Rows.Add(row3);
        //dtgv2.Rows.Add(row4);
        //dtgv2.Rows.Add(row5);
        //dtgv2.Rows.Add(row6);
        //dgv2.AutoGenerateColumns = false;
        //dgv2.DataSource = dtgv2;
        //}


        private void BindGrid1()
        {
            DataTable dtgv1 = new DataTable();
            dtgv1.Columns.Add("MallID");
            //dtgv1.Columns.Add("EmailAdd");
            string[] row1 = new string[] { "2" };
            string[] row2 = new string[] { "1" };
            string[] row3 = new string[] { "2" };
            string[] row4 = new string[] { "4" };

            dtgv1.Rows.Add(row1);
            dtgv1.Rows.Add(row2);
            dtgv1.Rows.Add(row3);
            dtgv1.Rows.Add(row4);

            dgv1.AutoGenerateColumns = false;
            dgv1.DataSource = dtgv1;
        }
    }
}
