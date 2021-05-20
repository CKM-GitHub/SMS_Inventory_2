using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using System.Globalization;

namespace CKM_Controls
{
    public partial class TestBoxTime : TextBox
    {
        [Browsable(true)]
        [Category("Extended Properties")] 
        [Description("Textbox Type")]
        [DisplayName("CKM_Type")]
         public Type CKM_Type { get; set; }
        public enum Type
        {
            Normal = 0,
            Date = 1,
            Integer = 2,
            Price = 3,
            Hiragana = 4,
            Time=5
        }

        [Browsable(true)]
        [Category("Extended Properties")]
        [Description("To Check Required Field On Enter KeyDown")]
        [DisplayName("CKM_IsRequired")]
        public Type CKM_Reqired { get; set; }

        public TestBoxTime()
        {
            InitializeComponent();
        }

        public TestBoxTime(IContainer container)
        {
            container.Add(this);

            InitializeComponent();
            this.KeyDown+=TestBoxTime_KeyDown;
            this.Validating += CKM_TextBox_Validating;
        }

        private void TestBoxTime_KeyDown(object sender, KeyEventArgs e)         
        {
             if (e.KeyCode == Keys.Enter)
            {
                if (CKM_Type == Type.Time)
                {
                    if (!string.IsNullOrWhiteSpace(this.Text))
                    {
                        //int hr = 0, min = 0, sec = 0;
                        string hr = string.Empty, min = string.Empty, sec = string.Empty;
                        
                        if (IsInteger(this.Text.Replace(":", "")))
                        {
                            if (this.Text.Contains(":"))
                            {
                                //if (timestr.Length > 2)
                                //    sec = timestr[timestr.Length - 0].PadLeft(2, '0');
                                //else
                                //    sec = "00";
                                //hr = timestr[timestr.Length - 2].PadLeft(2,'0');
                                //min = timestr[timestr.Length - 1].PadLeft(2, '0');

                                string[] timestr = this.Text.Split(':');
                                if (timestr.Length > 2)
                                {
                                    sec = timestr[timestr.Length - 1].PadLeft(2, '0');
                                    min = timestr[timestr.Length - 2].PadLeft(2, '0');
                                    hr = timestr[timestr.Length - 3].PadLeft(2, '0');
                                }
                                else
                                {
                                    min = timestr[timestr.Length - 1].PadLeft(2, '0');
                                    hr = timestr[timestr.Length - 2].PadLeft(2, '0');
                                }
                              
                                this.Text = this.Text.Replace(":", "");
                            }
                            else if (this.Text.Length <= 2)
                                min = this.Text;

                            else if (this.Text.Length >= 5)
                            {
                                string time = this.Text;
                                time = time.PadLeft(6,'0');
                                sec = time.Substring(time.Length - 2);
                                min = time.Substring(time.Length - 4).Substring(0, 2);
                                hr = Convert.ToInt32(time.Substring(0, time.Length - 4)).ToString();

                            }
                            else
                            {
                                string time = this.Text;
                                time = time.PadLeft(4,'0');
                                min = time.Substring(time.Length - 2);
                                hr = Convert.ToInt32(time.Substring(0, time.Length - 2)).ToString();
                            }

                            //if (sec == "00")
                            //{
                            //    sec = string.Empty;
                            //}
                            //if (hr == "00")
                            //{
                            //    hr = string.Empty;
                            //}

                            if (string.IsNullOrWhiteSpace(sec))
                            {
                                sec = DateTime.Now.Second.ToString().PadLeft(2, '0');
                            }
                            if (string.IsNullOrWhiteSpace(hr))
                            {
                                hr = DateTime.Now.Hour.ToString();
                            }

                            string txt = hr.ToString().PadLeft(2,'0') + ":" + min.ToString().PadLeft(2,'0') + ":" + sec.ToString().PadLeft(2,'0');
                            if(CheckDate(txt))
                                this.Text = txt;
                            else
                            {
                                MessageBox.Show("Invalid Format", "Date Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                this.Focus();
                            }
                        }
                        else
                        {
                            MessageBox.Show("Plz,Enter Integer Numbers!", "Date Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            this.Focus();
                        }
                    }
                }
            }
        }

        private void CKM_TextBox_Validating(object sender, CancelEventArgs e)
        {
            if (CKM_Type == Type.Time)
            {
                if (!string.IsNullOrWhiteSpace(this.Text))
                {
                    if (IsInteger(this.Text.Replace(":", "")))
                    {
                        string hr = string.Empty, min = string.Empty, sec = string.Empty;
                        if (this.Text.Contains(":"))
                        {
                            string[] timestr = this.Text.Split(':');
                            if (timestr.Length > 2)
                            {
                                sec = timestr[timestr.Length - 1].PadLeft(2, '0');
                                min = timestr[timestr.Length - 2].PadLeft(2, '0');
                                hr = timestr[timestr.Length - 3].PadLeft(2, '0');
                            }
                            else
                            {
                                min = timestr[timestr.Length - 1].PadLeft(2, '0');
                                hr = timestr[timestr.Length - 2].PadLeft(2, '0');
                            }

                            this.Text = this.Text.Replace(":", "");
                        }
                        else if (this.Text.Length <= 2)
                            min = this.Text;

                        else if (this.Text.Length >= 5)
                        {
                            string time = this.Text;
                            time = time.PadLeft(6, '0');
                            sec = time.Substring(time.Length - 2);
                            min = time.Substring(time.Length - 4).Substring(0, 2);
                            hr = Convert.ToInt32(time.Substring(0, time.Length - 4)).ToString();
                        }
                        else
                        {
                            string time = this.Text;
                            time = time.PadLeft(4, '0');
                            min = time.Substring(time.Length - 2);
                            hr = Convert.ToInt32(time.Substring(0, time.Length - 2)).ToString();
                        }

                        if (string.IsNullOrWhiteSpace(sec))
                        {
                            sec = DateTime.Now.Second.ToString().PadLeft(2, '0');
                        }
                        if (string.IsNullOrWhiteSpace(hr))
                        {
                            hr = DateTime.Now.Hour.ToString();
                        }
                        string txt = hr.ToString().PadLeft(2, '0') + ":" + min.ToString().PadLeft(2, '0') + ":" + sec.ToString().PadLeft(2, '0');
                        if (CheckDate(txt))
                            this.Text = txt;
                        else
                        {
                            MessageBox.Show("Invalid Format", "Date Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            this.Focus();
                        }
                    }
                    else
                    {
                        MessageBox.Show("Plz,Enter Integer Numbers!", "Date Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        this.Focus();
                    }
                }
            }
        }

        private bool CheckDate(string value)
        {

            DateTime d;
            return DateTime.TryParseExact(value,
                       "HH:mm:ss",
                       System.Globalization.CultureInfo.InvariantCulture,
                       DateTimeStyles.None,
                       out d);
            
        }

        private bool IsInteger(string value)
        {
            int Num;
            if (Int32.TryParse(value, out Num))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}
 