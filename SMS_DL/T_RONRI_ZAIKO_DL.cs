using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SMS_DL
{
    public class T_RONRI_ZAIKO_DL : Base_DL
    {
        public bool T_RONRI_ZAIKO_Insert(string dateTime,string shopFlag)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@DateTime", dateTime);
            dic.Add("@ShopFlag", shopFlag);
            return InsertUpdateDeleteData(dic, "T_RONRI_ZAIKO_Insert");
        }
    }
}
