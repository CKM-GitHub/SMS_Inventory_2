using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace SMS_Entity
{
    public class M_MakerZaiko_H_Entity : Base_Entity
    {
        public string PatternCD { get; set; }
        public string PatternName { get; set; }
        public string InsertOperator { get; set; }
        public string InsertDateTime { get; set; }
        public string UpdateOperator { get; set; }
        public string UpdateDateTime { get; set; }
    }
}
