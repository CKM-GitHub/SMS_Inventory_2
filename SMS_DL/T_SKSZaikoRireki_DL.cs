using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SMS_DL
{
    public class T_SKSZaikoRireki_DL : Base_DL  
    {
        public DataTable T_SKSZaikoRireki_Select()
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            return SelectData(dic, "T_SKSZaikoRireki_Select");
        }

        public DataTable T_SKSZaikoRireki_DetailSelect(string RenkeiDateTime)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@RenkeiDateTime", RenkeiDateTime);

            return SelectData(dic, "T_SKSZaikoRireki_DetailSelect");
        }

        public DataTable ResendInventoryToSKS(string RenkeiDateTime)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@RenkeiDateTime", RenkeiDateTime);
            return SelectData(dic, "T_ResendInventorySKSRireki_Select");
        }

      
    }
}
