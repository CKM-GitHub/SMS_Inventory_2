using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Globalization;

namespace SMS.CustomControls
{
    public class DateTextBox : MaskedTextBox
    {
        CommonFunctions cf = new CommonFunctions();
        public string MaskFormat { get; set; }
        public DateTextBox()
        {
            //this.Mask = string.IsNullOrWhiteSpace(Format) ? "0000/00/00" : Format;
            this.PromptChar = ' ';
            this.KeyPress += DateTextBox_KeyPress;
            this.Leave += new EventHandler(DateTextBox_Leave);
        }

        private void DateTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (IsEmptyDate())
                if ((string.IsNullOrWhiteSpace(this.Text.ToString()) || (this.Mask == "0000/00/00")) && (!this.TextAlign.Equals("Right")))
                {
                    this.Focus();
                    this.Select(0, 0);
                }
        }

        private void DateTextBox_Leave(object sender, EventArgs e)
        {
            try
            {
                if(!IsEmptyDate())
                    if(!string.IsNullOrWhiteSpace(this.Text.ToString().Replace("-","").Replace("/","")))
                        DateTime.ParseExact(this.Text.ToString().Replace("-","/"), MaskFormat, CultureInfo.InvariantCulture);
            }
            catch(Exception)
            {
                //MessageBox.Show("Invalid Date", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                cf.DSP_MSG("E103",string.Empty, string.Empty, string.Empty, string.Empty, string.Empty);
                this.Focus();
                this.SelectAll();               
            }
        }

        public bool IsEmptyDate()
        {
            if(string.IsNullOrWhiteSpace(this.Text.ToString().Replace("/","").Replace("-","")))
                return true;
            return false;
        }       
    }
}
