using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SMS_Entity
{
    public class T_SugorakuInport_Entity : Base_Entity
    {
        public string InportStartTime { get; set; }
        public string InportEndTime { get; set; }
        public string InportStartKBN { get; set; }
        public string InportEndKBN { get; set; }
    }
}
