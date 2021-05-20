using SMS_Entity;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SMS_Entity;
using SMS_DL;

namespace SMS_BL
{
    public class PSKS0114C_BL : Base_BL
    {
        M_MultiPorpose_DL mmdl;
        M_MultiPorpose_Entity mme;
        L_Log_DL lldl;
        public PSKS0114C_BL()
        {
            mmdl = new M_MultiPorpose_DL();
            lldl = new L_Log_DL();
        }

        public M_MultiPorpose_Entity M_MultiPorpose_Num_Select()
        {
            mme = new M_MultiPorpose_Entity();
            DataTable dt = mmdl.DynamicSelectData("Num1", "M_MultiPorpose", "Where ID=108 And [Key]='1'");
            if (dt.Rows.Count > 0)
            {
                mme.Num1 = dt.Rows[0]["Num1"].ToString();
                return mme;
            }

            return mme;
        }

        public bool M_MultiPorposeNum1_Update(M_MultiPorpose_Entity mme, L_Log_Entity LData)
        {
            mmdl.StartTransaction();
            if (mmdl.M_MultiPorpose_Num1_Update(mme))
            {
                lldl.Transaction = mmdl.Transaction;
                if (lldl.L_Log_Insert(LData))
                {
                    mmdl.CommitTransaction();
                    return true;
                }
            }
            return false;
        }
    }
}
