using SMS_Entity;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SMS_DL;

namespace SMS_BL
{
    public class m_sports_BL : Base_BL
    {
        m_sports_DL msdl;
        public m_sports_BL()
        {
            msdl = new m_sports_DL();
        }

        public DataTable m_sports_Select(m_sports_Entity mse)
        {
            return msdl.m_sports_Select(mse);
        }
    }
}
