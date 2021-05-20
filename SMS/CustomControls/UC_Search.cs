using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SMS.PopUp;
using SMS_BL;
using SMS_Entity;


namespace SMS.CustomControls
{
    public partial class UC_Search : UserControl
    {
        Login_Info loginInfo = new Login_Info();
        Base_BL basebl = new Base_BL();
        CommonFunctions cf = new CommonFunctions();

        #region Properties
        private int flag;
        public int UC_Flag
        {
            get { return flag; }
            set { flag = value; } 
        }

        public enum Type { Pattern = 0, Maker = 1, Brand = 2, Sports = 3, Category = 4 , PCMN0101K = 5 , Nothing = 6 , MultiJANCD=7 , ITEM=8 , SKUCD=9 , MakerShohinCD=10, JANCD=11 }
        public Type UC_Type { get; set; }

        public string UC_Code
        {
            get { return txtCode.Text; }
            set { txtCode.Text = value; }
        }

        public string UC_Name
        {
            get { return lblName.Text; }
            set{lblName.Text = value;}
        }

        public bool UC_NameVisible
        {
            get { return lblName.Visible; }
            set { lblName.Visible = value; }
        }

        public int UC_Code_Width
        {
            get { return txtCode.Width; }
            set
            {
                txtCode.Width = value;
                btnSearch.Location = new Point(txtCode.Location.X + txtCode.Width, btnSearch.Location.Y);
                lblName.Location = new Point(btnSearch.Location.X + btnSearch.Width, lblName.Location.Y);
            }
        }

        public int UC_Name_Width
        {
            get { return lblName.Width; }
            set
            {
                lblName.Width = value;
            }
        }

        public bool UC_SearchEnable
        {
            get { return btnSearch.Enabled; }
            set { btnSearch.Enabled = value; }
        }

        public bool UC_IsRequired { get; set; }

        public bool UC_Exists { get; set; }
        #endregion

        /// <summary>Constructor
        /// 
        /// </summary>
        public UC_Search()
        {
            InitializeComponent();
        }

        /// <summary> keypress event
        /// handle which key should allow by UC_Type
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void txtCode_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (this.UC_Type == Type.MultiJANCD)
            {
                if (((e.KeyChar < 48 || e.KeyChar > 57) && e.KeyChar != 8 && e.KeyChar != 45))
                {
                    e.Handled = true;
                    if (UC_Flag == 5)
                    {
                        if (e.KeyChar == ',')
                            e.Handled = false;
                    }
                }
            }
            else if (this.UC_Type == Type.Brand || this.UC_Type == Type.JANCD || this.UC_Type == Type.Category || this.UC_Type == Type.Sports || this.UC_Type == Type.Pattern || this.UC_Type == Type.Maker)
            {
                if (((e.KeyChar < 48 || e.KeyChar > 57) && e.KeyChar != 8 && e.KeyChar != 45))
                {
                    e.Handled = true;
                }
                if ((e.KeyChar == '-') || ((e.KeyChar == '.' && ((sender as TextBox).Text.IndexOf('.') > -1))))
                {
                    e.Handled = true;
                }
            }
           
        }

        /// <summary>handle copy/pastte, F9
        /// Decide Copy Paste allow by UC_Type
        /// Call Search Form on F9 click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void txtCode_KeyUp(object sender, KeyEventArgs e)
        {
            if (UC_Type == Type.MultiJANCD || UC_Type == Type.JANCD || UC_Type == Type.Brand)
            {
                if (e.KeyData == (Keys.Control | Keys.V))
                {
                    (sender as TextBox).Paste();
                }
                else if (e.KeyData == (Keys.Control | Keys.C))
                {
                    (sender as TextBox).Copy();
                }
                else if (e.KeyData == (Keys.Control | Keys.A))
                {
                    (sender as TextBox).SelectAll();
                }
                else if (e.KeyData == (Keys.Control | Keys.X))
                {
                    (sender as TextBox).Cut();
                }
            }
            else if (e.KeyCode == Keys.F9)
                ShowSearchForm();
        }

        /// <summary>clear name textbox
        /// clear name textbox if code textbox is null
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void txtCode_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (UC_IsRequired)
                {
                    #region***BugNo1***
                    if (string.IsNullOrWhiteSpace(UC_Code))
                    {
                        cf.DSP_MSG("E102", "名前", string.Empty, string.Empty, string.Empty, string.Empty);
                        this.SetFocus();
                        return;
                    }
                    #endregion
                }

                if (txtCode.Text == string.Empty)
                    lblName.Text = string.Empty;
                else
                {
                    if (!SelectData())
                    {
                        cf.DSP_MSG("E101", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                    }
                    else
                        base.OnKeyDown(e);
                }
            }
        }

        /// <summary>determine maxlength by UC_Type
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void txtCode_Enter(object sender, EventArgs e)
        {
            if (this.UC_Type == Type.Pattern)
                txtCode.MaxLength = 3;
            else if (this.UC_Type == Type.Maker)
                txtCode.MaxLength = 6;
            else if (this.UC_Type == Type.Brand)
                txtCode.MaxLength = 5;
            else if (this.UC_Type == Type.Sports)
                txtCode.MaxLength = 3;
            else if (this.UC_Type == Type.Category)
                txtCode.MaxLength = 3;
            else if (this.UC_Type == Type.MakerShohinCD)
                txtCode.MaxLength = 30;
            else if (this.UC_Type == Type.ITEM)
                txtCode.MaxLength = 32;
            else if (this.UC_Type == Type.SKUCD)
                txtCode.MaxLength = 32;
            else if (this.UC_Type == Type.MultiJANCD)
                txtCode.MaxLength = 153;
            else if (this.UC_Type == Type.JANCD)
                txtCode.MaxLength = 13;
        }

        /// <summary>serach button click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSearch_Click(object sender, EventArgs e)
        {
            ShowSearchForm();
        }

        /// <summary>call search form by form type
        /// 
        /// </summary>
        public void ShowSearchForm()
        {
            if (this.UC_Type == Type.Pattern)
            {
                frmPattern_List patternList = new frmPattern_List();
                patternList.ShowDialog();

                if (!string.IsNullOrWhiteSpace(patternList.PatternCD))
                {
                    txtCode.Text = patternList.PatternCD;
                    lblName.Text = patternList.PatternName;
                    txtCode.Focus();
                }
            }
            else if (this.UC_Type == Type.Maker)
            {
                frmMaker_List makerList = new frmMaker_List();
                makerList.ShowDialog();

                if (!string.IsNullOrWhiteSpace(makerList.MakerCD))
                {
                    txtCode.Text = makerList.MakerCD;
                    lblName.Text = makerList.MakerName;
                    txtCode.Focus();
                }
            }
            else if (this.UC_Type == Type.Brand)
            {
                frmBrand_List brandList = new frmBrand_List();
                brandList.ShowDialog();

                if (!string.IsNullOrWhiteSpace(brandList.brandCD))
                {
                    txtCode.Text = brandList.brandCD;
                    lblName.Text = brandList.brandName;
                    txtCode.Focus();
                }
            }
            else if (this.UC_Type == Type.Sports)
            {
                frmM_Sports_List sportslist = new frmM_Sports_List();
                sportslist.ShowDialog();

                if(!string.IsNullOrWhiteSpace(sportslist.nc_sports))
                {
                    txtCode.Text = sportslist.nc_sports;
                    lblName.Text = sportslist.vm_sports;
                    txtCode.Focus();
                }
            }
            else if (this.UC_Type == Type.Category)
            {
                frmM_Bunrui_List bunruilist = new frmM_Bunrui_List();
                bunruilist.ShowDialog();

                if (!string.IsNullOrWhiteSpace(bunruilist.nc_bunrui))
                {
                    txtCode.Text = bunruilist.nc_bunrui;
                    lblName.Text = bunruilist.vm_bunrui;
                    txtCode.Focus();
                }
            }
            else if ((this.UC_Type == Type.PCMN0101K) || (this.UC_Type == Type.MakerShohinCD) || (this.UC_Type == Type.MultiJANCD) || (this.UC_Type == Type.JANCD) || (this.UC_Type == Type.ITEM) || (this.UC_Type == Type.SKUCD))
            {
                Form frm = this.ParentForm as Form;
                if (frm.Controls.Find("lblSOperator", true).Count() > 0)
                {
                    Label lbl = frm.Controls.Find("lblSOperator", true)[0] as Label;

                    //frmPCMN0101K pcmn0101k = new frmPCMN0101K();
                    loginInfo.OperatorName = lbl.Text;
                    frmPCMN0101K pcmn0101k = new frmPCMN0101K(loginInfo, UC_Flag, false);
                    pcmn0101k.ShowDialog();

                    if (!string.IsNullOrWhiteSpace(pcmn0101k.result))
                    {
                        txtCode.Text = pcmn0101k.result;
                        txtCode.Focus();
                    }
                }
            }
        }

        /// <summary>select data from relative table
        /// 
        /// </summary>
        /// <returns></returns>
        public bool SelectData()
        {
            string fields = string.Empty;
            string tableName = string.Empty;
            string condition = string.Empty;
            DataTable dtResult = new DataTable();
            if (this.UC_Type == Type.Pattern)
            {
                fields = "PatternName";
                tableName = "M_MakerZaiko_H";
                condition = "Where PatternCD = '" + txtCode.Text + "'";

                dtResult = basebl.DynamicSelectData(fields, tableName, condition);

            }
            else if (this.UC_Type == Type.Maker)
            {
                fields = "VM_SHIIRESAKI";
                tableName = "M_SHIIRESAKI";
                condition = "Where VC_SHIIRESAKI = '" + txtCode.Text + "'";

                dtResult = basebl.DynamicSelectData(fields, tableName, condition);
            }
            else if (this.UC_Type == Type.Brand)
            {
                fields = "vm_brand";
                tableName = "m_brand";
                condition = "Where nc_brand = '" + txtCode.Text + "'";

                dtResult = basebl.DynamicSelectData(fields, tableName, condition);
            }
            else if (this.UC_Type == Type.Category)
            {
                fields = "vm_bunrui";
                tableName = "m_bunrui";
                condition = "Where nc_bunrui = '" + txtCode.Text + "'";

                dtResult = basebl.DynamicSelectData(fields, tableName, condition);
            }
            else if (this.UC_Type == Type.Sports)
            {
                fields = "vm_sports";
                tableName = "m_sports";
                condition = "Where nc_sports = '" + txtCode.Text + "'";

                dtResult = basebl.DynamicSelectData(fields, tableName, condition);
            }
            else if (this.UC_Type == Type.ITEM)
            {
                fields = "vc_user_shohin";
                tableName = " t_hanbai_shohin ";
                condition = "where vc_user_shohin= '" + txtCode.Text + "'";

                dtResult = basebl.DynamicSelectData(fields, tableName, condition);
            }
            else if (this.UC_Type == Type.MultiJANCD)
            {
                fields = "VC_JAN1";
                tableName = "M_KIHON_SHOHIN";
                string jan = "'" + txtCode.Text.Replace(",", "','") + "'";
                condition = "where VC_JAN1 IN(" + jan + ")";

                dtResult = basebl.DynamicSelectData(fields, tableName, condition);
            }
            else if (this.UC_Type == Type.MakerShohinCD)
            {
                fields = "vc_maker_shohin";
                tableName = "t_kihon_shohin";
                condition = "where vc_maker_shohin= '" + txtCode.Text + "'";

                dtResult = basebl.DynamicSelectData(fields, tableName, condition);
            }
            else if (this.UC_Type == Type.SKUCD)
            {
                fields = "HanbaiShohinCD";
                tableName = "T_MakerZaiko";
                condition = " where HanbaiShohinCD = '" + txtCode.Text + "'";

                dtResult = basebl.DynamicSelectData(fields, tableName, condition);

            }
            if (dtResult.Rows.Count > 0)
            {
                lblName.Text = dtResult.Rows[0][fields].ToString();
                UC_Exists = true;
                return true;
            }
            else
                lblName.Text = string.Empty;
            UC_Exists = false;
            return false;
        }

        /// <summary>set focus tot code textbox
        /// 
        /// </summary>
        public void SetFocus()
        {
            txtCode.Focus();
        }

        /// <summary>disable control
        /// code and button disable
        /// </summary>
        public void Disabled()
        {
            txtCode.Enabled = btnSearch.Enabled = false;
        }

        /// <summary>enable control
        /// code and button enable
        /// </summary>
        public void Enable()
        {
            txtCode.Enabled = btnSearch.Enabled = true;
        }

        /// <summary>clear data
        /// 
        /// </summary>
        public void Clear()
        {
            UC_Code = UC_Name = string.Empty;
        }
    }
}
