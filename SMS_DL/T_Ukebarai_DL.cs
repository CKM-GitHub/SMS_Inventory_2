using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SMS_DL
{
    public class T_Ukebarai_DL : Base_DL
    {
        public DataTable T_Ukebarai_Select()
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            return SelectData(dic, "T_Ukebarai_Select");
        }

        public bool PSKS0106B_Insert(string CSV_xml)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@xml_csv", CSV_xml);

            return InsertUpdateDeleteData(dic, "PSKS0106B_Insert");
        }
    }
}
