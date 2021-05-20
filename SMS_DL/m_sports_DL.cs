using SMS_Entity;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SMS_DL
{
    public class m_sports_DL : Base_DL
    {
        public DataTable m_sports_Select(m_sports_Entity mse)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@vm_sports", mse.vm_sports);
            dic.Add("@nc_sports", mse.nc_sports);

            return SelectData(dic, "m_sports_Select");
        }
    }
}
