using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SMS_Entity;
using SMS_DL;

namespace SMS_BL
{
    public class PSKS0110S_BL : Base_BL
    {
        ItemSearch_DL isdl;
        public PSKS0110S_BL()
        {
            isdl = new ItemSearch_DL();
        }

        public DataTable PSKS0110S_Select(ItemSearch_Entity ise)
        {
            return isdl.PSKS0110S_Select(ise);
        }

        public DataTable PSKS0110S_Search_Select(ItemSearch_Entity ise)
        {
            return isdl.PSKS0110S_Search_Select(ise);
        }

        public DataTable GetFolder()
        {
            return isdl.DynamicSelectData("Char3,Char2", "M_MultiPorpose", "WHERE ID = '110' AND [Key] = '1'");
        }
    }
}
