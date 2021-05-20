using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SMS_Entity
{
    public class L_Log_Entity : Base_Entity
    {
        public string SEQ { get; set; }
        public string OperateDate { get; set; }
        public string OperateTime { get; set; }
        public string InsertOperator { get; set; }
        public string Program { get; set; }
        public string PC { get; set; }
        public string OperateMode { get; set; }
        public string KeyItem { get; set; }
    }
}
