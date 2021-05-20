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
    public class m_bunrui_BL
    {
        m_bunrui_DL mbdl;
        public m_bunrui_BL()
        {
            mbdl = new m_bunrui_DL();
        }

        public DataTable m_bunrui_Select(m_bunrui_Entity mbe)
        {
            return mbdl.m_bunrui_Select(mbe);
        }
    }
}
