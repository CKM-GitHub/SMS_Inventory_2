using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SMS_DL;
using SMS_Entity;
using System.Data;

namespace SMS_BL
{
    public class PETC0303I_BL
    {
        T_NyuukinMeisai_DL NyuukinMeisaiDL;
        T_Tessuryou_DL TessuryouDL;
        M_MultiPorpose_DL mdl;
        M_ShiwakeControl_DL msdl;
        L_Log_DL LogDL;
        public PETC0303I_BL()
        {
            NyuukinMeisaiDL = new T_NyuukinMeisai_DL();
            TessuryouDL = new T_Tessuryou_DL();
            mdl = new M_MultiPorpose_DL();
            msdl = new M_ShiwakeControl_DL();
            LogDL = new L_Log_DL();
        }

        public DataSet PETC0303I_SELECT(string opcd)
        {
            return TessuryouDL.PETC0303I_SELECT(opcd);
        }

        public bool PETC0303IDelete(L_Log_Entity log_data)
        {
            return NyuukinMeisaiDL.PETC0303IDelete(log_data);
          
        }
        public bool CheckExistName1_0303I(T_NyuukinMeisai_Entity NyuukinMeisai_data, L_Log_Entity log_data, string twobytes)
        {
            if (NyuukinMeisaiDL.CheckExistName1_0303I(NyuukinMeisai_data, log_data, twobytes).Rows.Count > 0)
            {
                return true;
            }
            return false;
        }
        public bool PETC0303I_Insert(T_NyuukinMeisai_Entity NyuukinMeisai_data, L_Log_Entity log_data,string twobytes)
        {
            if (NyuukinMeisaiDL.PETC0303I_Insert(NyuukinMeisai_data, log_data, twobytes))
            {

                     return true;

            }
            return false;
        }
    }
}
