using SMS_DL;
using SMS_Entity;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SMS_BL
{
    public class ZaikoHenkan_BL : Base_BL
    {
        T_Zaiko_DL tzdl;
        L_Log_DL lldl;
        public ZaikoHenkan_BL()
        {
            tzdl = new T_Zaiko_DL();
            lldl = new L_Log_DL();
        }

        public bool T_Zaiko_INSERT_UPDATE(T_Zaiko_Entity tze,L_Log_Entity lle)
        {
             if(tzdl.T_Zaiko_INSERT_UPDATE(tze))
             {
                 return lldl.L_Log_Insert(lle);
             }
             return false;
        }
    }
}
