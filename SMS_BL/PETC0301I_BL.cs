using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using SMS_DL;
using SMS_Entity;

namespace SMS_BL
{
    public class PETC0301I_BL : Base_BL
    {
        M_MultiPorpose_DL mdl;
        T_JUCHU_DL juchuDL;

        public PETC0301I_BL()
        {
            mdl = new M_MultiPorpose_DL();
            juchuDL = new T_JUCHU_DL();
        }

        public DataTable YahooShop_Combo()
        {
            return mdl.DynamicSelectData("[Key],Char1", "M_MultiPorpose", "WHERE ID = '111' ORDER BY [Key]"); 
        }

        public DataTable GetFolder(string ShopID)
        {
            return mdl.DynamicSelectData("Char3,Char2", "M_MultiPorpose", "WHERE ID = '111' AND [Key] = '" + ShopID + "'");
        }

        public DataTable PECT0301I_Select(string ShopID, string startDate, string endDate)
        {
            return juchuDL.PECT0301I_Select(ShopID, startDate, endDate);
        }
    }
}
