using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using SMS_Entity;

namespace SMS_DL
{
    public class T_JUCHU_DL : Base_DL
    {
        public DataTable PECT0301I_Select(string ShopID, string startDate, string endDate)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            dic.Add("@ShopID",ShopID);
            dic.Add("@startDate", startDate);
            dic.Add("@endDate", endDate);

            return SelectData(dic, "PECT0301I_Select");
        }

        public DataTable PECT0307I_Select(string ShopID, string startDate, string endDate)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            dic.Add("@ShopID", ShopID);
            dic.Add("@startDate", startDate);
            dic.Add("@endDate", endDate);

            return SelectData(dic, "PECT0307I_Select");
        }
    }
}
