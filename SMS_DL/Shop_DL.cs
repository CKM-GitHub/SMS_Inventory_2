using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SMS_DL
{
    public class Shop_DL : Base_DL
    {
        public bool ImportShopData(string shopFlag)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@ShopFlag", shopFlag);
            return InsertUpdateDeleteData(dic, "PSKS0113B_Import");
        }
    }
}
