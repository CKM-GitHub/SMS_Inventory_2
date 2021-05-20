using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CKM_Controls
{
    public partial class CKM_Label : UserControl
    {
        public string value { get { return lblCKM.Text; } set { lblCKM.Text = value; } }

        public Color ChangeLableColor
        {
            get { return lblCKM.BackColor;}
            set { lblCKM.BackColor = value; }
        }

        public ContentAlignment TextAlign 
        {
            get { return lblCKM.TextAlign; }
            set { lblCKM.TextAlign = value; } 
        }

        public CKM_Label()
        {
            InitializeComponent();
            //this.Font = new Font(this.Font, FontStyle.Bold);
            this.Font = new Font("MS Gothic", 9, FontStyle.Bold);
            this.BackColor = Color.Transparent;
        }
    }
}
