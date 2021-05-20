using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SMS_DL;
using System.Data;
using SMS_Entity;

namespace SMS_BL
{
    public class PSKS0120S_BL : Base_BL
    {
        M_Zaiko_Kanri_DL zaikokanridl;
        ItemSearch_DL Isdl;
        L_Log_DL logdl;
        
        public PSKS0120S_BL()
        {
            zaikokanridl = new M_Zaiko_Kanri_DL();
            Isdl = new ItemSearch_DL();
            logdl = new L_Log_DL();
        }

        public DataTable PSKS0120S_InitialSelect(M_Zaiko_Kanri_Entity zaikokanri)
        {
            return zaikokanridl.PSKS0120S_InitialSelect(zaikokanri);
        }

        public DataTable PSKS0120S_Select(ItemSearch_Entity ise,string OperatorCD)
        {
            return Isdl.PSKS0120S_Select(ise, OperatorCD);
        }

        public DataTable GetFolder()
        {
            return Isdl.DynamicSelectData("Char3,Char2", "M_MultiPorpose", "WHERE ID = '110' AND [Key] = '2'");
        }

        public bool PSKS0120S_Update(M_Zaiko_Kanri_Entity zaikokanri,L_Log_Entity lle)
        {
            zaikokanri.UpdateDateTime = lle.OperateDate = System.DateTime.Now.ToString();
            zaikokanridl.StartTransaction();
            logdl.Transaction = zaikokanridl.Transaction;

            if (zaikokanridl.PSKS0120S_Update(zaikokanri))
            {
                if(logdl.L_Log_Insert(lle))
                {
                    zaikokanridl.CommitTransaction();
                    return true;
                }
            }
            return false;
        }
    }
}
