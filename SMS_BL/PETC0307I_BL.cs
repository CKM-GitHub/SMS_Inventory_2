using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SMS_BL;
using SMS_DL;
using System.Data;

namespace SMS_BL
{
   public class PETC0307I_BL:Base_BL
    {
        M_MultiPorpose_DL mdl;
        T_JUCHU_DL juchuDL;

        public PETC0307I_BL()
        {
            mdl = new M_MultiPorpose_DL();
            juchuDL = new T_JUCHU_DL();
        }

        public DataTable RakutenShop_Combo()
        {
            return mdl.DynamicSelectData("[Key],Char1", "M_MultiPorpose", "WHERE ID = '115' ORDER BY [Key]");
        }

        public DataTable PECT0301I_Select(string ShopID, string startDate, string endDate)
        {
            return juchuDL.PECT0307I_Select(ShopID, startDate, endDate);
        }

        public DataTable GetFolder(string ShopID)
        {
            return mdl.DynamicSelectData("Char3,Char2", "M_MultiPorpose", "WHERE ID = '115' AND [Key] = '" + ShopID + "'");
        }
    }
}
