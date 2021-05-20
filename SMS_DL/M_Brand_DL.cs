using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using SMS_Entity;

namespace SMS_DL
{
    public class M_Brand_DL : Base_DL
    {
        public DataTable M_Brand_Select(M_Brand_Entity m_Brand_Data)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            dic.Add("@BrandCode", m_Brand_Data.BrandCD);
            dic.Add("@BrandName", m_Brand_Data.BrandName);

            return SelectData(dic, "M_Brand_Select");
        }
    }
}
