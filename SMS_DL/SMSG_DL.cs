using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using SMS_Entity;

namespace SMS_DL
{
    public class SMSG_DL : Base_DL
    {
        public DataTable SMSG_Select(SMSG_Entity smsg_data)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@MSGSB", smsg_data.MSGSB);
            dic.Add("@MSGID", smsg_data.MSGID);

            return SelectData(dic, "SMSG_Select");
        }
        public bool INSERT_ErrorTable(string error, string formName)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@ErrorMsg", error);
            dic.Add("@FormName", formName);
            return InsertUpdateDeleteData(dic, "INSERT_ErrorTable");

        }
    }
}
