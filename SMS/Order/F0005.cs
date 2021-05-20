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

namespace SMS.Order
{
    public partial class frmF0005 : SMS_Base
    {
        Login_Info loginInfo;

        public frmF0005()
        {
            InitializeComponent();
        }

        public frmF0005(Login_Info loginInfo)
            : base(loginInfo)
        {
            InitializeComponent();
            this.loginInfo = loginInfo;
            SetDesignerFunction();
        }

        private void SetDesignerFunction()
        {
            FunctionButtonDisabled(6);
            FunctionButtonDisabled(11);
            FunctionButtonDisabled(10);

            dgvF0005.DisabledColumn("colNo,colMakerShohinCD,colSKUCD,colName,colColor,colSize,colTeika,colStockDate,colHatchuNote,colUnit");
            FormMode = "登録";
        }

        private void F0005_Load(object sender, EventArgs e)
        {
            FormName = "発注入力";
            //BindGrid();
            BindCombo();
        }

        private void BindGrid()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("No");
            dt.Columns.Add("JANCD");
            dt.Columns.Add("MakerShohinCD");
            dt.Columns.Add("SKUCD");
            dt.Columns.Add("Name");
            dt.Columns.Add("Color");
            dt.Columns.Add("Size");
            dt.Columns.Add("Teika");
            dt.Columns.Add("Kakeritsu");
            dt.Columns.Add("StockDate");
            dt.Columns.Add("InternalComment");
            dt.Columns.Add("ExternalComment");
            dt.Columns.Add("HatchuNote");
            dt.Columns.Add("HatchuSu");
            dt.Columns.Add("Unit");
            dt.Columns.Add("HatchuTanka");
            dt.Columns.Add("HatchuGaku");

            string[] row1 = new string[] { "1", "1234567890123", "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3", "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸ15", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸ15", 
                                           "99,999,999", "999.99", "9999/99/99", "社内備考・コメント10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "社外備考・コメント10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", 
                                           "-99,999","ＸＸＸ","99,999,999", "-999,999,999"};
            string[] row2 = new string[] { "2", "1234567890123", "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3", "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸ15", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸ15", 
                                           "99,999,999", "999.99", "9999/99/99", "社内備考・コメント10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "社外備考・コメント10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", 
                                           "-99,999","ＸＸＸ","99,999,999", "-999,999,999"};
            string[] row3 = new string[] { "3", "1234567890123", "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3", "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸ15", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸ15", 
                                           "99,999,999", "999.99", "9999/99/99", "社内備考・コメント10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "社外備考・コメント10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", 
                                           "-99,999","ＸＸＸ","99,999,999", "-999,999,999"};
            string[] row4 = new string[] { "4", "1234567890123", "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3", "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸ15", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸ15", 
                                           "99,999,999", "999.99", "9999/99/99", "社内備考・コメント10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "社外備考・コメント10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", 
                                           "-99,999","ＸＸＸ","99,999,999", "-999,999,999"};
            string[] row5 = new string[] { "5", "1234567890123", "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3", "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸ15", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸ15", 
                                           "99,999,999", "999.99", "9999/99/99", "社内備考・コメント10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "社外備考・コメント10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", 
                                           "-99,999","ＸＸＸ","99,999,999", "-999,999,999"};
            string[] row6 = new string[] { "6", "1234567890123", "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3", "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸ15", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸ15", 
                                           "99,999,999", "999.99", "9999/99/99", "社内備考・コメント10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "社外備考・コメント10ＸＸＸＸＸＸＸＸＸ20ＸＸＸＸＸＸＸＸＸ30ＸＸＸＸＸＸＸＸＸ40", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20", 
                                           "-99,999","ＸＸＸ","99,999,999", "-999,999,999"};

            dt.Rows.Add(row1);
            dt.Rows.Add(row2);
            dt.Rows.Add(row3);
            dt.Rows.Add(row4);
            dt.Rows.Add(row5);
            dt.Rows.Add(row6);

            dgvF0005.AutoGenerateColumns = false;
            dgvF0005.DataSource = dt;
        }

        private void BindCombo()
        {
            #region 店舗

            DataTable dtstore = new DataTable();

            dtstore.Columns.Add("storeID");
            dtstore.Columns.Add("storeName");

            dtstore.Rows.Add("-1", "");
            dtstore.Rows.Add("1", "Web店");
            dtstore.Rows.Add("2", "豊中店");
            dtstore.Rows.Add("3", "石橋店");
            dtstore.Rows.Add("4", "三宮店");
            dtstore.Rows.Add("5", "江坂店");
           
            dtstore.AcceptChanges();
            comboStore.SelectedIndex = -1;

            comboStore.DisplayMember = "storeName";
            comboStore.ValueMember = "storeID";


            this.comboStore.DataSource = dtstore;

            #endregion

            #region 自社倉庫

            DataTable dt2 = new DataTable();

            dt2.Columns.Add("warehouseID");
            dt2.Columns.Add("warehouseName");

            dt2.Rows.Add("-1", "");
            dt2.Rows.Add("1", "本社Web倉庫");
            dt2.Rows.Add("1", "豊中店倉庫");
            dt2.Rows.Add("1", "石橋店倉庫");
            dt2.Rows.Add("1", "三宮店倉庫");
            dt2.Rows.Add("1", "江坂店倉庫");

            dt2.AcceptChanges();
            
            comboBox2.SelectedIndex = -1;

            comboBox2.DisplayMember = "warehouseName";
            comboBox2.ValueMember = "warehouseID";

            this.comboBox2.DataSource = dt2;
            #endregion



        }

     
    }
}
