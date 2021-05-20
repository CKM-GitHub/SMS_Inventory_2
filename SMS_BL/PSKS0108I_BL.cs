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
    public class PSKS0108I_BL : Base_BL
    {
        T_SugorakuInport_DL tsidl;
        M_MultiPorpose_DL mmdl;
        T_RONRI_ZAIKO_DL trzdl;
        L_Log_DL lldl;
        public PSKS0108I_BL()
        {
            tsidl = new T_SugorakuInport_DL();
            mmdl = new M_MultiPorpose_DL();
            trzdl = new T_RONRI_ZAIKO_DL();
            lldl = new L_Log_DL();
        }

        public DataTable T_SugorakuInport_SELECT()
        {
            return tsidl.T_SugorakuInport_SELECT();
        }

        public bool btnWeb_Update(T_SugorakuInport_Entity tsie, M_MultiPorpose_Entity mme)
        {
            tsidl.StartTransaction();
            if (tsidl.T_SugorakuInport_Update(tsie))
            {
                mme.ID = "108";
                if (mmdl.M_MultiPorpose_Num1_Update(mme))
                {
                    mme.ID = "109";
                    if (mmdl.M_MultiPorpose_Num1_Update(mme))
                    {
                        tsidl.CommitTransaction();
                        return true;
                    }
                }
            }
            return false;
        }

        public bool T_SugorakuInport_Update(T_SugorakuInport_Entity tsie)
        {
            return tsidl.T_SugorakuInport_Update(tsie);
        }

        public bool T_RONRI_ZAIKO_Insert(L_Log_Entity LogData, string dateTime,string shopFlag)
        {
            trzdl.StartTransaction();
            if (trzdl.T_RONRI_ZAIKO_Insert(dateTime,shopFlag))
            {
                lldl.Transaction = trzdl.Transaction;
                lldl.L_Log_Insert(LogData);
                trzdl.CommitTransaction();
                return true;
            }
            return false;
        }

        public bool M_MultiPorpose_Num1_Update(M_MultiPorpose_Entity mme)
        {
            return mmdl.M_MultiPorpose_Num1_Update(mme);
        }
    }
}
