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
    public partial class F0017 : SMS_Base
    {
        Login_Info loginInfo;

        public F0017()
        {
            InitializeComponent();
        }

        public F0017(Login_Info loginInfo)
            : base(loginInfo)
        {
            InitializeComponent();
            this.loginInfo = loginInfo;
            SetDesignerFunction();
        }

        private void SetDesignerFunction()
        {
            lblMode.Visible = false;
            FunctionButtonDisabled(2);
            FunctionButtonDisabled(3);
            FunctionButtonDisabled(4);
            FunctionButtonDisabled(6);
            FunctionButtonDisabled(7);
            FunctionButtonDisabled(8);
            FunctionButtonDisabled(9);
            FunctionButtonDisabled(10);
            FunctionButtonDisabled(11);
        }

        private void F0017_Load(object sender, EventArgs e)
        {
            FormName = "出荷指示登録（受注分）";
            dgv.DisabledColumn("Column1,Column3,Column5,Column6,Column7,Column8,Column9,Column10,Column11,Column12,Column13,Column14,Column15");//Column16,Column17,Column18,Column1,Column2,Column1,Column2,Column1,Column2,Column1,Column2,Column1,Column2,
            SetDesignerFunction();
            //BindGrid();
        }

        //private void BindGrid()
        //{
        //    DataTable dtgrid = new DataTable();

        //    dtgrid.Columns.Add("No.");
        //    dtgrid.Columns.Add("ShippingDate");
        //    dtgrid.Columns.Add("ShippingNum");
        //    dtgrid.Columns.Add("DesignatedDay");
        //    dtgrid.Columns.Add("OrderNumber");
        //    dtgrid.Columns.Add("SKUCD");
        //    dtgrid.Columns.Add("JANCD");
        //    dtgrid.Columns.Add("ProductName");
        //    dtgrid.Columns.Add("ColorSizeRemark");
        //    dtgrid.Columns.Add("DeliveryCompany");
        //    dtgrid.Columns.Add("Warehouse");
        //    dtgrid.Columns.Add("ShipmentQuantity");
        //    dtgrid.Columns.Add("ShipToPartyName");
        //    dtgrid.Columns.Add("ShipmentRemarks");
        //    dtgrid.Columns.Add("ShipmentRemarks(internal)");

        //    string[] row1 = new string[] { "1", "9999/99/99","1234567890", "9999/99/99 99", "12345678901", "XXXXXXXXX1XXXXXXXX2XXXXXXXX3", "XXXXXXXXX1XX3", "ＸＸＸＸＸＸＸＸＸＸＸＸＸＸ40", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20","ＸＸＸＸＸＸＸＸＸ10Ｘ", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸ15", "999,999", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸ15", "出荷備考(社外)ＸＸ10ＸＸＸＸＸＸＸＸＸＸＸＸ40", "出荷備考(社内)ＸＸ10ＸＸＸＸＸＸＸＸＸＸＸＸＸ40"};
        //    string[] row2 = new string[] { "2", "9999/99/99","1234567890", "9999/99/99 99", "12345678901", "XXXXXXXXX1XXXXXXXX2XXXXXXXX3", "XXXXXXXXX1XX3", "ＸＸＸＸＸＸＸＸＸＸＸＸＸＸ40", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20","ＸＸＸＸＸＸＸＸＸ10Ｘ", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸ15", "999,999", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸ15", "出荷備考(社外)ＸＸ10ＸＸＸＸＸＸＸＸＸＸＸＸ40", "出荷備考(社内)ＸＸ10ＸＸＸＸＸＸＸＸＸＸＸＸＸ40" };
        //    string[] row3 = new string[] { "3", "9999/99/99","1234567890", "9999/99/99 99", "12345678901", "XXXXXXXXX1XXXXXXXX2XXXXXXXX3", "XXXXXXXXX1XX3", "ＸＸＸＸＸＸＸＸＸＸＸＸＸＸ40", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20","ＸＸＸＸＸＸＸＸＸ10Ｘ", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸ15", "999,999", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸ15", "出荷備考(社外)ＸＸ10ＸＸＸＸＸＸＸＸＸＸＸＸ40", "出荷備考(社内)ＸＸ10ＸＸＸＸＸＸＸＸＸＸＸＸＸ40" };
        //    string[] row4 = new string[] { "4", "9999/99/99","1234567890", "9999/99/99 99", "12345678901", "XXXXXXXXX1XXXXXXXX2XXXXXXXX3", "XXXXXXXXX1XX3", "ＸＸＸＸＸＸＸＸＸＸＸＸＸＸ40", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20","ＸＸＸＸＸＸＸＸＸ10Ｘ", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸ15", "999,999", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸ15", "出荷備考(社外)ＸＸ10ＸＸＸＸＸＸＸＸＸＸＸＸ40", "出荷備考(社内)ＸＸ10ＸＸＸＸＸＸＸＸＸＸＸＸＸ40" };
        //    string[] row5 = new string[] { "5", "9999/99/99","1234567890", "9999/99/99 99", "12345678901", "XXXXXXXXX1XXXXXXXX2XXXXXXXX3", "XXXXXXXXX1XX3", "ＸＸＸＸＸＸＸＸＸＸＸＸＸＸ40", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20","ＸＸＸＸＸＸＸＸＸ10Ｘ", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸ15", "999,999", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸ15", "出荷備考(社外)ＸＸ10ＸＸＸＸＸＸＸＸＸＸＸＸ40", "出荷備考(社内)ＸＸ10ＸＸＸＸＸＸＸＸＸＸＸＸＸ40" };
        //    string[] row6 = new string[] { "6", "9999/99/99","1234567890", "9999/99/99 99", "12345678901", "XXXXXXXXX1XXXXXXXX2XXXXXXXX3", "XXXXXXXXX1XX3", "ＸＸＸＸＸＸＸＸＸＸＸＸＸＸ40", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸＸＸＸＸＸ20","ＸＸＸＸＸＸＸＸＸ10Ｘ", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸ15", "999,999", "ＸＸＸＸＸＸＸＸＸ10ＸＸＸＸ15", "出荷備考(社外)ＸＸ10ＸＸＸＸＸＸＸＸＸＸＸＸ40", "出荷備考(社内)ＸＸ10ＸＸＸＸＸＸＸＸＸＸＸＸＸ40" };

        //    dtgrid.Rows.Add(row1);
        //    dtgrid.Rows.Add(row2);
        //    dtgrid.Rows.Add(row3);
        //    dtgrid.Rows.Add(row4);
        //    dtgrid.Rows.Add(row5);
        //    dtgrid.Rows.Add(row6);

        //    dgv.AutoGenerateColumns = false;
        //    dgv.DataSource = dtgrid;

        //}

        private void btnAllCheck_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                dgv[1, i].Value = true;//check every row in the forth column

            }
        }

        private void btnAllCancel_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                dgv[1, i].Value = false;//uncheck every row in the forth column

            }
        }
    }
}
