using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using SMS_Entity;

namespace SMS_DL
{
    public class Menu_DL : Base_DL
    {
        public DataTable Menu_Select(Login_Info loginInfo)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@KAICD", loginInfo.CompanyCode);
            dic.Add("@TMPCD", loginInfo.StoreCode);
            dic.Add("@BMNCD", loginInfo.DepCode);
            dic.Add("@TANCD", loginInfo.OperatorCode);

            return SelectData(dic, "Menu_SELECT");
        }
    }
}
