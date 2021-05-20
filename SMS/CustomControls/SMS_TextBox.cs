using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;

namespace SMS.CustomControls
{
    public class SMS_TextBox : TextBox
    {
        [Browsable(true)]
        [Category("Extended Properties")]
        [Description("txt_Type")]
        [DisplayName("txt_Type")]
        public Type txt_Type { get; set; }    
        public enum Type
        {
            Normal = 0, 
            IntegerOnly = 1,
            Hiragana = 2,
            Decimal = 3
        }

        byte DecimalPlace = 0;      
        [Browsable(true)]
        [Category("Extended Properties")]
        [Description("Decimal Place")]
        [DisplayName("DecimalPlace")]
        public byte Decimalplace
        {
            get { return DecimalPlace; }
            set { DecimalPlace = value; }
        }

        bool useThousandSeparator = true;
        [Browsable(true)]
        [Category("Extended Properties")]
        [Description("UseThousandSeparator")]
        [DisplayName("UseThousandSeparator")]
        public bool UseThousandSeparator
        {
            get { return useThousandSeparator; }
            set { useThousandSeparator = value; }
        }

        bool useMinus = false;
        [Browsable(true)]
        [Category("Extended Properties")]
        [Description("UseMinus")]
        [DisplayName("UseMinus")]
        public bool UseMinus
        {
            get { return useMinus; }
            set { useMinus = value; }
        }

        public SMS_TextBox()
        {
            this.Enter += txt_Enter;
            this.KeyPress += txt_KeyPress;
            this.KeyUp += txt_KeyUp;
            this.TextChanged += new EventHandler(SMS_TextBox_TextChanged);
            this.Leave += txt_Leave;
        }


        private System.Drawing.Color _backColorDisabled = System.Drawing.Color.LightGray;// System.Drawing.Color.Gainsboro;
        private System.Drawing.Color _foreColorDisabled = System.Drawing.SystemColors.ControlText;
        private void txt_Enter(object sender, EventArgs e)
        {
            foreach (InputLanguage lang in InputLanguage.InstalledInputLanguages)
            {
                if (lang.LayoutName.Equals("Japanese"))
                {
                    InputLanguage.CurrentInputLanguage = lang;
                    break;
                }
            }

            this.BackColor = Color.LightYellow;
            if (this.txt_Type == Type.Hiragana)
                this.ImeMode = System.Windows.Forms.ImeMode.Hiragana;
            if (this.txt_Type == Type.Decimal)
                this.TextAlign = HorizontalAlignment.Right;
        }

        private void txt_Leave(object sender, EventArgs e)
        {
            this.BackColor = Color.White;
        }

        private void txt_KeyPress(object sender, KeyPressEventArgs e)      
        {
            //if (string.IsNullOrWhiteSpace(this.Text) && (!this.TextAlign.Equals("Right")))
            //{
                //this.Focus();
                //this.Select(0, 0);

                if (this.txt_Type == Type.IntegerOnly)
                {
                    if (((e.KeyChar < 48 || e.KeyChar > 57) && e.KeyChar != 8 && e.KeyChar != 46 && e.KeyChar != 45))
                    {
                        e.Handled = true;
                    }
                }
                if (this.txt_Type == Type.Decimal)
                {
                    if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
                    {
                        e.Handled = true;
                    }

                    // only allow one decimal point
                    if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
                    {
                        e.Handled = true;
                    }
                }
            //}
        }

        private void txt_KeyUp(object sender, KeyEventArgs e)
        {
            if (this.txt_Type == Type.IntegerOnly || this.txt_Type == Type.Decimal)
            
            {
               if(e.KeyData == (Keys.Control | Keys.V))
                {
                    (sender as TextBox).Paste();
                }
               else if (e.KeyData == (Keys.Control | Keys.C))
               {
                   (sender as TextBox).Copy();
               }
            }
        }

        public void Enable()
        {
            this.Enabled = true;
            this.BackColor = Color.White;
            this.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
        }

        public void Disabled()
        {
            this.Enabled = false;
            this.ForeColor = Color.Black;
            this.BackColor = Color.FromArgb(217, 217, 217);
            this.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
        }

        protected override void OnEnabledChanged(EventArgs e)
        {
            base.OnEnabledChanged(e);
            //if (!(this.Enabled))
            //{
            //    this.SetStyle(System.Windows.Forms.ControlStyles.UserPaint, true);
            //}
            //else
            //{
            //    this.SetStyle(System.Windows.Forms.ControlStyles.UserPaint, false);
            //}

            //this.Invalidate();
        }

        protected override void OnPaint( System.Windows.Forms.PaintEventArgs e )
        {
            base.OnPaint( e );
            //System.Drawing.SolidBrush textBrush;

            ////Set user selected backcolor when disabled
            //if (this.Enabled)
            //{
            //    textBrush = new System.Drawing.SolidBrush( this.ForeColor );
            //}
            //else
            //{
            //    System.Drawing.Color backColorDisabled = this._backColorDisabled;
            //    if (this.Parent.FindForm() != null)
            //    {
            //        backColorDisabled = this.Parent.FindForm().BackColor;
            //    }
            //    textBrush = new System.Drawing.SolidBrush( this._foreColorDisabled );
            //System.Drawing.SolidBrush backBrush = new System.Drawing.SolidBrush(backColorDisabled);
            //e.Graphics.FillRectangle( backBrush, ClientRectangle );
            //this.BackColor = Color.FromArgb(217, 217, 217);
            //this.BorderStyle = System.Windows.Forms.BorderStyle.None;
            //    Pen pen = new Pen(Color.Black);
            //}

            ////Paint text
            //e.Graphics.DrawString(this.Text, this.Font, textBrush, 1.0F, 1.0F);

        }

        private void SMS_TextBox_TextChanged(object sender, EventArgs e)
        {
            if (this.txt_Type == Type.Decimal &&  useThousandSeparator)
            {
                TextBox tb = sender as TextBox;
                string value = tb.Text.Replace(",", "");
                ulong ul;
                if (ulong.TryParse(value, out ul))
                {
                    tb.TextChanged -= SMS_TextBox_TextChanged;
                    if(!tb.Text.Equals("0"))
                        tb.Text = string.Format("{0:#,#}", ul);
                    tb.SelectionStart = tb.Text.Length;
                    tb.TextChanged += SMS_TextBox_TextChanged;
                }

                //TextBox txt = sender as TextBox;
                //string num = txt.Text;
                //string dec = string.Empty;
                //string minus = string.Empty;

                //if (txt.Text.Contains("."))
                //{
                //    string[] str = txt.Text.Split('.');
                //    num = str[0];
                //    dec += "." + str[1];
                //}

                //if (txt.Text.Contains("-"))
                //{
                //    minus = "-";
                //}

                //string value = num.Replace(",", "").Replace("-", "");
                //ulong ul;
                //if (ulong.TryParse(value, out ul))
                //{
                //    txt.TextChanged -= SMS_TextBox_TextChanged;
                //    int start = txt.SelectionStart;
                //    int oldCount = txt.Text.Split(',').Length - 1;//comma count
                //    //txt.Text = string.Format("{0:#,#}", ul);//format thousand seperator    //att delete for zero value display
                //    txt.Text += dec;
                //    txt.Text = minus + txt.Text;
                //    int newCount = txt.Text.Split(',').Length - 1;//comma count
                //    if (!string.IsNullOrWhiteSpace(txt.Text))
                //    {
                //        if (oldCount < newCount)
                //            txt.SelectionStart = start + 1;
                //        else if (oldCount > newCount)
                //            txt.SelectionStart = start - 1 > 0 ? start - 1 : 0;
                //        else
                //            txt.SelectionStart = start;
                //    }
                //    else txt.SelectionStart = 0;

                //    txt.TextChanged += SMS_TextBox_TextChanged;
                //}
            }
        }
    }
}
