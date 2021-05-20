using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace SMS_Entity
{
    public class MakerZaiko_D_Entity : Base_Entity
    {
        public string PatternCD { get; set; }
        public string SEQ { get; set; }
        public string ItemName { get; set; }
        public string ItemProperty { get; set; }
        public DataTable dtKamoKu { get; set; }
        public string KamokuXML { get; set; }
    }
}
