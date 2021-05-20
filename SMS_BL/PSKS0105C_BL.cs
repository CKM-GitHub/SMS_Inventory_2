using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SMS_DL;
using SMS_Entity;

namespace SMS_BL
{
    public class PSKS0105C_BL : Base_BL
    {
        M_MultiPorpose_DL mmdl;
        T_Ukebarai_DL tudl;
        L_Log_DL lldl;
        public PSKS0105C_BL()
        {
            mmdl = new M_MultiPorpose_DL();
            lldl = new L_Log_DL();
            tudl = new T_Ukebarai_DL();
        }

        public string M_MultiPorpose_Num3_Select()
        {
            DataTable dt = mmdl.DynamicSelectData("Num3", "M_MultiPorpose", "Where ID=102 AND [Key]='1'");
            if (dt.Rows.Count > 0)
            {
                return dt.Rows[0]["Num3"].ToString();
            }

            return string.Empty;
        }

        public DataTable M_MultiPorpose_MultiSelect()
        {
            return mmdl.DynamicSelectData("*", "M_MultiPorpose", "where ID=102 OR ID=103 OR ID=107 AND [Key]='1'");
        }

        public bool M_MultiPorpose_Num1_Update(M_MultiPorpose_Entity mme,L_Log_Entity lle)
        {
            mmdl.StartTransaction();
            if (mmdl.M_MultiPorpose_Num1_Update(mme))
            {
                lldl.Transaction = mmdl.Transaction;
                if (lldl.L_Log_Insert(lle))
                {
                    mmdl.CommitTransaction();
                    return true;
                }
            }
            return false;
        }

        public DataTable T_Ukebarai_Select()
        {
            return tudl.T_Ukebarai_Select();
        }
    }
}
