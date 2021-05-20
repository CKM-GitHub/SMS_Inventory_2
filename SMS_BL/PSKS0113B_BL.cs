using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SMS_DL;
using System.Data;
using SMS_Entity;

namespace SMS_BL
{
    public class PSKS0113B_BL :Base_BL
    {
        Shop_DL shopDL = new Shop_DL();

        public M_MultiPorpose_Entity M_MultiPurpose_Select()
        {
            DataTable dt = DynamicSelectData("Num1,Num2", "M_MultiPorpose", "where ID= 107 AND [Key]='1'");
            M_MultiPorpose_Entity mme = new M_MultiPorpose_Entity();
            if (dt.Rows.Count > 0)
            {
                mme.Num1 = dt.Rows[0]["Num1"].ToString();
                mme.Num2 = dt.Rows[0]["Num2"].ToString();
            }

            return mme;
        }

        public bool ImportShopData(string shopFlag)
        {
            return shopDL.ImportShopData(shopFlag);
        }
    }
}
