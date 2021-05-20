using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SMS_Entity;
using System.Data;
using System.Data.SqlClient;

namespace SMS_DL
{
    public class SMNU_DL : Base_DL
    {
        public DataTable SMNU_SELECT(MOPE_Entity mopeData)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@KAICD", mopeData.KAICD);
            dic.Add("@KNGCD", mopeData.KNGCD);
            dic.Add("@MNUCD", mopeData.MNUCD);

            return SelectData(dic, "SMNU_SELECT");
        }
    }
}
