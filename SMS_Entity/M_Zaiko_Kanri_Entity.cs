using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SMS_Entity
{
    public class M_Zaiko_Kanri_Entity : Base_Entity
    {
        public string OperatorCD { get; set; }
        public string AdminCD { get; set; }
        public string HanbaiShohinCD { get; set; }
        public string JanCD { get; set; }
        public string vc_maker_shohin { get; set; }
        public string taisho { get; set; }
        public string kijunsu { get; set; }
        public string Zaiko_keisan { get; set; }
        public string shijisho { get; set; }
        public string Memo { get; set; }
        public string InsertDateTime { get; set; }
        public string UpdateDateTime { get; set; }

        public DataTable dt { get; set; }

        public string xml { get; set; }
    }
}
