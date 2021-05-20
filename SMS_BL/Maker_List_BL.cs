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
    public class Maker_List_BL : Base_BL //Maker_List
    {
        M_SHIIRESAKI_DL M_Shiiresaki_DL = new M_SHIIRESAKI_DL();

        public DataTable Maker_Select(M_SHIIRESAKI_Entity M_Shiiresaki_Data)
        {
            return M_Shiiresaki_DL.M_Shiiresaki_Select(M_Shiiresaki_Data);
        }
    }
}
