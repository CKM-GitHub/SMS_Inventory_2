using SMS_Entity;
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
    public class PSKS0111C_BL : Base_BL
    {
        M_MultiPorpose_DL mmdl;
        T_SKSZaikoRireki_DL tszrdl;
        L_Log_DL lldl;
        Base_DL bdl;

        public PSKS0111C_BL()
        {
            mmdl = new M_MultiPorpose_DL();
            tszrdl = new T_SKSZaikoRireki_DL();
            lldl = new L_Log_DL();
        }

        public M_MultiPorpose_Entity M_MultiPorpose_Num_Select()
        {
            M_MultiPorpose_Entity mme = new SMS_Entity.M_MultiPorpose_Entity();
            DataTable dt = mmdl.DynamicSelectData("Num1,Num3", "M_MultiPorpose", "Where ID=109 And [Key]='1'");
            if (dt.Rows.Count > 0)
            {
                mme.Num1 = dt.Rows[0]["Num1"].ToString();
                mme.Num3 = dt.Rows[0]["Num3"].ToString();
            }

            return mme;
        }

        public bool M_MultiPorpose_Num1_Update(M_MultiPorpose_Entity mme, L_Log_Entity lle)
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

        public DataTable T_SKSZaikoRireki_Select()
        {
            return tszrdl.T_SKSZaikoRireki_Select();
        }

        public DataTable T_SKSZaikoRireki_DetailSelect(string RenkeiDateTime)
        {
            return tszrdl.T_SKSZaikoRireki_DetailSelect(RenkeiDateTime);
        }

        public DataTable ResendInventoryToSKS(string RenkeiDateTime)
        {
            return tszrdl.ResendInventoryToSKS(RenkeiDateTime);
        }

    }
}
