using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace SMS_Entity
{
   public class SuuryoHenKan_Entity
    {
        public string PatternCD { get; set; }
        public string SEQ { get; set; }
        public string Suuryo { get; set; }
        public string HenkanSuuryo { get; set; }
        public string InsertOperator { get; set; }
        public string InsertDateTime { get; set; }
        public string UpdateOperator { get; set; }
        public string UpdateDateTime { get; set; }
        public DataTable dtSuuryo { get; set; }
        public string SuuryoXML { get; set; }
    }
}
