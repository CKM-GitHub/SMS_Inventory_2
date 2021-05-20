using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using SMS_Entity;

namespace SMS_DL 
{
    public class M_SHIIRESAKI_DL : Base_DL
    {
        public DataTable M_Shiiresaki_Select(M_SHIIRESAKI_Entity M_Shiiresaki_Entity) //Pattern_List_BL
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            dic.Add("@VC_SHIIRESAKI", M_Shiiresaki_Entity.MakerCD);
            dic.Add("@VM_SHIIRESAKI", M_Shiiresaki_Entity.MakerName);

            return SelectData(dic, "M_Shiiresaki_Select");
        }
    }
}
