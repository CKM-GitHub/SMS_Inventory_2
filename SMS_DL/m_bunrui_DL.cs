using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SMS_Entity;

namespace SMS_DL
{
    public class m_bunrui_DL : Base_DL
    {
        public DataTable m_bunrui_Select(m_bunrui_Entity mbe)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@nc_bunrui", mbe.nc_bunrui);
            dic.Add("@vm_bunrui", mbe.vm_bunrui);
            return SelectData(dic, "m_bunrui_Select");
        }
    }
}
