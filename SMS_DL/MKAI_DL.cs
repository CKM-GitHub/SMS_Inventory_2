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
    public class MKAI_DL : Base_DL 
    {
        public DataTable MKAI_SELECT(MKAI_Entity mkai_data) // Login_BL
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@KAICD", mkai_data.KAICD);
            return SelectData(dic, "MKAI_SELECT");
        }

    }
}
