using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;

namespace SMS.CustomControls
{
    public class SMS_Button : Button
    {
        public SMS_Button()
        {
            this.FlatStyle = FlatStyle.Popup;
            this.BackColor = Color.FromArgb(191,191,191);
            this.Cursor = Cursors.Hand;
        }
    }
}
