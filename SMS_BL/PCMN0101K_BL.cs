using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SMS_Entity;
using SMS_DL;
using System.Data;

namespace SMS_BL
{
    public class PCMN0101K_BL : Base_BL
    {
        ItemSearch_DL isdl;
        public PCMN0101K_BL()
        {
            isdl = new ItemSearch_DL();
        }

        public DataTable PCMN0101K_Search(ItemSearch_Entity ise)
        {
            return isdl.PCMN0101K_Search(ise);
        }
    }
}
