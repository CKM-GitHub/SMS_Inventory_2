using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SMS_Entity;
using System.Data;
using SMS_DL;
namespace SMS_BL
{
    public class PETC0310I_BL
    {
        T_NyuukinMeisai_DL NyuukinMeisaiDL = new T_NyuukinMeisai_DL();
        T_Tessuryou_DL TessuryouDL = new T_Tessuryou_DL();
        M_MultiPorpose_DL mdl = new M_MultiPorpose_DL();
        T_SmileShiwake_DL msdl = new T_SmileShiwake_DL();
        L_Log_DL LogDL = new L_Log_DL();
        public  PETC0310I_BL()
        {
           // return NyuukinMeisaiDL.DynamicSelectData("*", "M_ShiwakeControl_MF", "where [KEYCODE]= '1'");
        }
        public DataTable T_SmileHistory()
        {
            return msdl.T_SmileHistorySelect();
        }
        public DataTable GetTSmile()
        {
            return msdl.GetTSmile();
        }
        public bool T_SmileInsert(string opt, string xml)
        {
            return msdl.T_SmileInsert(opt,xml);
        }
        public DataTable T_SmileHistoryInsert_Output()
        {
            return msdl.T_SmileHistoryInsert_Output();
        }
        //T_SmileHistoryInsert

    }
}
