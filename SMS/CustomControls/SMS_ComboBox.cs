using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SMS_BL;
using System.Data;
using System.Drawing;

namespace SMS.CustomControls
{
    public class SMS_ComboBox : ComboBox
    {
        [Browsable(true)]
        [Category("Extended Properties")]
        [Description("tableName")]
        [DisplayName("Type")]
        public cbo_Type Cbo_Type { get; set; }
        public enum cbo_Type
        {
            m_nendo = 0,
            m_sports = 1,
            m_supplier = 2,
            m_bunrui = 3,
            m_season = 4,
            m_brand = 5,
            SpecialFlag = 6,
            M_SHIIRESAKI = 7,
            YoyakuFlag = 8,
            Taisho=9,
            ZaikoKeisan=10
        }

        Base_BL bbl;
        public SMS_ComboBox()
        {
            bbl = new Base_BL();
            this.AutoCompleteMode = AutoCompleteMode.Suggest;
            this.AutoCompleteSource = AutoCompleteSource.CustomSource;

            this.Enter += cbo_Enter;
            this.Leave += cbo_Leave;

            // Required for ownerdraw
            this.DrawItem += new DrawItemEventHandler(EnableDisplayCombo_DrawItem);
            //// Required for ownerdraw
            this.DrawMode = DrawMode.OwnerDrawFixed;
            this.EnabledChanged += new EventHandler(EnableDisplayCombo_EnabledChanged);
        }

        private void cbo_Enter(object sender, EventArgs e)
        {
            this.BackColor = Color.LightYellow;
        }

        private void cbo_Leave(object sender, EventArgs e)
        {
            this.BackColor = Color.White;
        }

        public void Bind()
        {
            if (Cbo_Type == cbo_Type.m_nendo)
            {
                DataTable dt = bbl.DynamicSelectData("nk_nendo,vm_nendo", "m_nendo", "order by nn_row");
                BindCombo("nk_nendo", "vm_nendo", dt);             
            }
            else if (Cbo_Type == cbo_Type.m_season)
            {
                DataTable dt = bbl.DynamicSelectData("nk_season,vm_season", "m_season", "order by nn_row");
                BindCombo("nk_season", "vm_season", dt);
            }
            else if (Cbo_Type == cbo_Type.m_brand)
            {
                DataTable dt = bbl.DynamicSelectData("nc_brand,vm_brand", "m_brand", "order by vm_brand");
                BindCombo("nc_brand", "vm_brand", dt);
            }
            else if (Cbo_Type == cbo_Type.m_sports)
            {
                DataTable dt = bbl.DynamicSelectData("nc_sports,vm_sports", "m_sports", "order by nn_row");
                BindCombo("nc_sports", "vm_sports", dt);
            }
            else if (Cbo_Type == cbo_Type.m_bunrui)
            {
                DataTable dt = bbl.DynamicSelectData("nc_bunrui,vm_bunrui", "m_bunrui", "order by nn_row");
                BindCombo("nc_bunrui", "vm_bunrui", dt);
            }
            else if (Cbo_Type == cbo_Type.M_SHIIRESAKI)
            {
                DataTable dt = bbl.DynamicSelectData("VC_SHIIRESAKI,VM_SHIIRESAKI", "M_SHIIRESAKI", "order by VM_SHIIRESAKI");
                BindCombo("VC_SHIIRESAKI", "VM_SHIIRESAKI", dt);
            }
            else if (Cbo_Type == cbo_Type.SpecialFlag)
            {
                DataTable dt = new DataTable();
                dt.Columns.Add("nf_maker_hachu_huka");
                dt.Columns.Add("name");
                dt.Rows.Add("1", "【★】");
                dt.Rows.Add("3", "【■】");
                dt.Rows.Add("5", "【F】");
                dt.Rows.Add("6", "【□】");
                dt.Rows.Add("8", "[送料別途]");

                BindCombo("nf_maker_hachu_huka", "name", dt);
            }
            else if (Cbo_Type == cbo_Type.YoyakuFlag)
            {
                DataTable dt = new DataTable();
                dt.Columns.Add("nk_yoyaku_shohin");
                dt.Columns.Add("name");
                dt.Rows.Add("1", "-");
                dt.Rows.Add("2", "予約");
                dt.Rows.Add("3", "☆即☆");
                dt.Rows.Add("4", "【B】");

                BindCombo("nk_yoyaku_shohin", "name", dt);
            }
            else if (Cbo_Type == cbo_Type.Taisho)
            {
                DataTable dt = new DataTable();
                dt.Columns.Add("taishoID");
                dt.Columns.Add("name");
                dt.Rows.Add("0", "無効");
                dt.Rows.Add("1", "有効");

                BindCombo("taishoID", "name", dt);
            }
            else if (Cbo_Type == cbo_Type.ZaikoKeisan)
            {
                DataTable dt = new DataTable();
                dt.Columns.Add("zaikoKeisan");
                dt.Columns.Add("name");
                dt.Rows.Add("0", "WEB在庫のみ");
                dt.Rows.Add("1", "豊中在庫数+Web通販在庫数");

                BindCombo("zaikoKeisan", "name", dt);
            }
        }

        private void BindCombo(string key,string value,DataTable dt)
        {
            DataRow dr = dt.NewRow();
            dr[key] = "-1";
            dt.Rows.InsertAt(dr, 0);
            this.DataSource = dt;
            this.DisplayMember = value;
            this.ValueMember = key;
        }

        // If control is disabled, we switch to DropDownList style, so we can control the appearance of the
        // edit box
        void EnableDisplayCombo_EnabledChanged(object sender, EventArgs e)
        {
            if (this.Enabled)
                this.DropDownStyle = ComboBoxStyle.DropDown;
            else
                this.DropDownStyle = ComboBoxStyle.DropDownList;
        }
        // Ownerdraw routine
        void EnableDisplayCombo_DrawItem(object sender, DrawItemEventArgs e)
        {
            System.Drawing.Graphics g = e.Graphics;
            Rectangle r = e.Bounds;

            if (e.Index >= 0)
            {
                DataRowView drv = this.Items[e.Index] as DataRowView;
                string label = string.Empty;;
                if (drv != null)
                    label = drv.Row[1].ToString();

                // This is how we draw a disabled control
                if (e.State == (DrawItemState.Disabled | DrawItemState.NoAccelerator | DrawItemState.NoFocusRect | DrawItemState.ComboBoxEdit))
                {
                    e.Graphics.FillRectangle(new SolidBrush(Color.FromArgb(169, 208, 142)), r);
                    g.DrawString(label, e.Font, Brushes.Black, r);
                    e.DrawFocusRectangle();
                }
                // This is how we draw the items in an enabled control that aren't in focus
                else if (e.State == (DrawItemState.NoAccelerator | DrawItemState.NoFocusRect))
                {
                    e.Graphics.FillRectangle(new SolidBrush(Color.White), r);
                    g.DrawString(label, e.Font, Brushes.Black, r);
                    e.DrawFocusRectangle();
                }
                // This is how we draw the focused items
                else if (e.State == (DrawItemState.NoAccelerator))
                {
                    e.Graphics.FillRectangle(new SolidBrush(Color.White), r);
                    g.DrawString(label, e.Font, Brushes.Black, r);
                    e.DrawFocusRectangle();
                }
                else
                {
                    e.Graphics.FillRectangle(new SolidBrush(Color.FromArgb(169, 208, 142)), r);
                    g.DrawString(label, e.Font, Brushes.Black, r);
                    e.DrawFocusRectangle();
                }


            }
            g.Dispose();
        }
    }
}
