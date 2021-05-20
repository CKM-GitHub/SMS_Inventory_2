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
    public class Brand_List_BL : Base_BL
    {
        M_Brand_DL m_Brand_DL = new M_Brand_DL();

        public DataTable Brand_Select(M_Brand_Entity m_Brand_Data)
        {
            return m_Brand_DL.M_Brand_Select(m_Brand_Data);
        }
    }
}
