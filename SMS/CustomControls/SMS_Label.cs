using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;

namespace SMS.CustomControls
{
    public class SMS_Label : Label
    {
        public SMS_Label() : base()
        {
            this.Font = new Font("MS Gothic",9, FontStyle.Bold);
            this.BackColor = Color.Transparent;
        }
    }
}
