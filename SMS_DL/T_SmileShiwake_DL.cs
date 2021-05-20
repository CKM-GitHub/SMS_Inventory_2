using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
namespace SMS_DL
{
   public class T_SmileShiwake_DL :Base_DL
    {

        public DataTable T_SmileHistorySelect()
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            return SelectData(dic, "T_SmileHistorySelect");
        }//
        public DataTable GetTSmile()
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            return SelectData(dic, "T_SmileShiwakeSelect");
        }
        public bool T_SmileInsert(string opt, string xml)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("opt",opt);
            dic.Add("xml", xml);
            return InsertUpdateDeleteData(dic, "T_SmileInsert");
        }
        public DataTable T_SmileHistoryInsert_Output()
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            return SelectData(dic, "T_SmileHistoryInsert_Output");
        }
        //

    }
}
