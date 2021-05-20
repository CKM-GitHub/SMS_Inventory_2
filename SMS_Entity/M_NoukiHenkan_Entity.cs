using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace SMS_Entity
{
    public class M_NoukiHenkan_Entity
    {
        public string PatternCD { get; set; }
        public string SEQ { get; set; }
        public string Nouki { get; set; }
        public string HenkanNouki { get; set; }
        public string InsertOperator { get; set; }
        public string InsertDateTime { get; set; }
        public string UpdateOperator { get; set; }
        public string UpdateDateTime { get; set; }
        public DataTable dtNouki { get; set; }
        public string NoukiXML { get; set; }
    }
}
