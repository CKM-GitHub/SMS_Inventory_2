using SMS_Entity;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SMS_DL
{
    public class T_SugorakuInport_DL : Base_DL
    {
        public DataTable T_SugorakuInport_SELECT()
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            return SelectData(dic, "T_SugorakuInport_SELECT");
        }

        public bool T_SugorakuInport_Update(T_SugorakuInport_Entity ts_data)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@InportStartTime", ts_data.InportStartTime);
            dic.Add("@InportStartKBN", ts_data.InportStartKBN);
            dic.Add("@InportEndTime", ts_data.InportEndTime);
            dic.Add("@InportEndKBN", ts_data.InportEndKBN);
            return InsertUpdateDeleteData(dic, "T_SugorakuInport_Update");
        }
    }
}
