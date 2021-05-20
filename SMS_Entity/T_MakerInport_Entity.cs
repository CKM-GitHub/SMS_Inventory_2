using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SMS_Entity
{
    public class T_MakerInport_Entity : Base_Entity
    {
        public string InportSEQ { get; set; }
        public string DataSourseMakerCD { get; set; }
        public string InportFile { get; set; }
        public string PatternCD { get; set; }
        public string InportSu { get; set; }
        public string Nouki { get; set; }
    }
}
