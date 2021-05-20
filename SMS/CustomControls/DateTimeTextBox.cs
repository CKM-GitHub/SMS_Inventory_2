using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Globalization;
using System.Drawing;
using SMS;

namespace SMS.CustomControls
{
    public class DateTimeTextBox : MaskedTextBox
    {
        public string MaskFormat { get; set; }
        public bool DateError = false;
        public DateTimeTextBox()
        {
            this.Mask = "0000/00/00 00:00:00";
            this.PromptChar = ' ';

            this.Leave += new EventHandler(DateTimeTextBox_Leave);
            this.KeyPress += DateTimeTextBox_KeyPress;
        }

        private void DateTimeTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (IsEmptyDateTime())
                if ((string.IsNullOrWhiteSpace(this.Text.ToString()) || (this.Mask == "0000/00/00 00:00:00")) && (!this.TextAlign.Equals("Right")))
                {
                    this.Focus();
                    this.Select(0, 0);
                }
        }

        private void DateTimeTextBox_Leave(object sender, EventArgs e)
        {
            try
            {
                if (!IsEmptyDateTime())
                {
                    if (!string.IsNullOrWhiteSpace(this.Text.Replace("-", "").Replace("/", "")))
                        DateTime.ParseExact(this.Text.Replace("-", "/"), MaskFormat, CultureInfo.InvariantCulture);
                    this.DateError = false;
                }
            }
            catch (Exception)
            {
                CommonFunctions cf = new CommonFunctions();
                cf.DSP_MSG("E113", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty); 
                this.Focus();
                this.DateError = true;
                this.SelectAll();
            }
        }

        public bool IsEmptyDateTime()
        {
            if (string.IsNullOrWhiteSpace(this.Text.Replace("/", "").Replace("-", "").Replace(":","")))
                return true;
            return false;
        }

        public bool IsEmptyDate()
        {
            throw new NotImplementedException();
        }
    }
}
