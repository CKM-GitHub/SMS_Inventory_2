using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SMS_DL;
using SMS_Entity;
using System.Data;

namespace SMS_DL
{
    public class PSKS0112B_DL : Base_DL
    {
        public DataTable PSKS0112B_Select(Base_Entity be)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@DateTime", be.InsertDateTime);

            return SelectData(dic, "PSKS0112B_Select");
        }
    }
}
