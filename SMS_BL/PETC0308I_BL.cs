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
   public class PETC0308I_BL : Base_BL
    {

         T_NyuukinMeisai_DL NyuukinMeisaiDL;
        T_Tessuryou_DL TessuryouDL;
        M_MultiPorpose_DL mdl;
        M_ShiwakeControl_DL msdl;
        L_Log_DL LogDL;

        public PETC0308I_BL()
        {
            NyuukinMeisaiDL = new T_NyuukinMeisai_DL();
            TessuryouDL = new T_Tessuryou_DL();
            mdl = new M_MultiPorpose_DL();
            msdl = new M_ShiwakeControl_DL();
            LogDL = new L_Log_DL();
        }

        public DataTable Mshiwake_title()
        {
            return NyuukinMeisaiDL.DynamicSelectData("*", "M_ShiwakeControl_MF", "where [KEYCODE]= '1'");
        }
        public DataTable GetFolder()
        {
            return mdl.DynamicSelectData("Char1,Char2", "M_MultiPorpose", "WHERE ID = '119' AND [Key] = 1");
        }

        public bool T_NyuukinMeisai_T_Tesuuryou_Delete(T_NyuukinMeisai_Entity NyuukinMeisai_data, T_Tesuuryou_Entity Tessuryou_data, L_Log_Entity log_data)
        {
            NyuukinMeisaiDL.StartTransaction();
            if (NyuukinMeisaiDL.T_NyuukinMeisai_Tesuuryou_Delete(log_data))
            {
                if (LogDL.L_Log_Insert(log_data))
                {
                    NyuukinMeisaiDL.CommitTransaction();

                    return true;
                }
            }
            return false;
        }

        public bool Check_Motonames(T_NyuukinMeisai_Entity NyuukinMeisai_data)
        {
            return NyuukinMeisaiDL.Check_Motonames(NyuukinMeisai_data).Rows.Count > 0 ? true : false;
        }
        public bool PETC0308I_Insert(T_NyuukinMeisai_Entity NyuukinMeisai_data, T_Tesuuryou_Entity Tessuryou_data, L_Log_Entity log_data)
        {
            //if (NyuukinMeisaiDL.PETC0308I_Insert(NyuukinMeisai_data, Tessuryou_data, log_data)) {
            //    return false;
            //}
            if (NyuukinMeisaiDL.T_NyuukinMeisai_Insert_MF(NyuukinMeisai_data))
            {
                if (TessuryouDL.T_Tesuuryou_Insert_MF(Tessuryou_data))
                {
                    if (LogDL.L_Log_Insert(log_data))
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        public DataTable PETC0308I_SELECT(T_Tesuuryou_Entity tse)
        {
            return TessuryouDL.PETC0302I_SELECT(tse);
        }
        public DataTable PETC0308I_SELECT_MF(T_Tesuuryou_Entity tse)
        {
            return TessuryouDL.PETC0308I_SELECT_MF(tse);
        }
        public DataTable GetTaxRate(string date)
        {
            return msdl.GetTaxRate_MF(date);
        }

        public DataTable PETC0308I_Detail_SELECT(T_NyuukinMeisai_Entity NyuukinMeisai_data)
        {
            return NyuukinMeisaiDL.PETC0308I_Detail_SELECT(NyuukinMeisai_data);
        }
        public bool Update_Tesuyou_Nyuukin(DataTable tesuyou, DataTable nyuukin)
        {
            return TessuryouDL.Update_Tesuyou_Nyuukin_MF(DataTableToXml(tesuyou), DataTableToXml(nyuukin));
        }
        public DataTable NyuukinDate_NyuukinMoto_Select(string opcd, string DataKBN)
        {
            // var fileds="select distinct  * from (   select  CONVERT(varchar(10),KouzaNyuukinDate,111)as KouzaNyuukinDate,NyuukinMotoCD,NyuukinMotoName from  T_NyuukinMeisai WHERE NouseFlag <> 1 Group By KouzaNyuukinDate , NyuukinMotoCD , NyuukinMotoName  union  All select kouzanyuukinDate, NyuukinMOtoCD, NyuukinMotoName2 from  T_Tesuuryou  WHERE NouseFlg <> 1 group by kouzanyuukinDate, NyuukinMOtoCD, NyuukinMotoName2 ) t";
            //var table_name="";
            //  var where_Con= "";
            return NyuukinMeisaiDL.DynamicSelectData("CONVERT(varchar(10),KouzaNyuukinDate,111)as KouzaNyuukinDate,NyuukinMotoCD,NyuukinMotoName", "T_NyuukinMeisai", "WHERE NouseFlag <> 1 and OrderPriority <> 0 and DataKBN = " + DataKBN + " and OperatorCD = '" + opcd + "' Group By KouzaNyuukinDate , NyuukinMotoCD , NyuukinMotoName  order by NyuukinMotoCD asc");
        }
        public DataTable PETC0302I_CSVExport1(T_NyuukinMeisai_Entity tnu, string exp)
        {

            return NyuukinMeisaiDL.PETC0302I_CSVExport1(tnu, exp);
        }
        public bool Check_Motonames_MF(T_NyuukinMeisai_Entity NyuukinMeisai_data)
        {
            return NyuukinMeisaiDL.Check_Motonames_MF(NyuukinMeisai_data).Rows.Count > 0 ? true : false;
        }

       //Check_Motonames_MF

    }
}
