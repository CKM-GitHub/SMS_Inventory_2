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
    public class MOPE_DL : Base_DL
    {
        public DataTable MOPE_SELECT(MOPE_Entity mopeData) //Login_BL
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@KAICD", mopeData.KAICD);
            dic.Add("@TANCD", mopeData.TANCD);

            return SelectData(dic, "MOPE_SELECT");
        }

        public bool MOPE_UPDATE(MOPE_Entity mopeData) //LoginPassword
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@TANCD", mopeData.TANCD);
            dic.Add("@PASWD", mopeData.PASWD);

            return InsertUpdateDeleteData(dic, "MOPE_UPDATE");
        }
    }
}
