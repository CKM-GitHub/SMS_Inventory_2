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
   public  class PETC0309I_BL :Base_BL
    {
        T_NyuukinMeisai_DL NyuukinMeisaiDL= new T_NyuukinMeisai_DL();
        T_Tessuryou_DL TessuryouDL= new T_Tessuryou_DL();
        M_MultiPorpose_DL mdl= new M_MultiPorpose_DL();
        M_ShiwakeControl_DL msdl= new M_ShiwakeControl_DL();
        L_Log_DL LogDL= new L_Log_DL();
       //public  PETC0309I_BL()
       //{
           //NyuukinMeisaiDL = new T_NyuukinMeisai_DL();
           //TessuryouDL = new T_Tessuryou_DL();
           //mdl = new M_MultiPorpose_DL();
           //msdl = new M_ShiwakeControl_DL();
           //LogDL = new L_Log_DL();
       //}
        //GetTaxRate
       public DataTable Mshiwake_title()
       {
           return NyuukinMeisaiDL.DynamicSelectData("*", "M_ShiwakeControl_MF", "where [KEYCODE]= '1'");
       }
       public DataTable GetTaxRate(string date)
       {
           return msdl.GetTaxRate_MF(date);
       }
       public DataSet PETC0309I_SELECT(string opcd)
       {
           return TessuryouDL.PETC0309I_SELECT(opcd);
       }

       public bool Update_Tesuyou_Nyuukin(DataTable tesuyou, DataTable nyuukin)
       {
           return TessuryouDL.Update_Tesuyou_Nyuukin_MF_309I(DataTableToXml(tesuyou), DataTableToXml(nyuukin));
       }
       public DataTable GetFolder3i()
       {
           return mdl.DynamicSelectData("Char1,Char2", "M_MultiPorpose", "WHERE ID = '120' AND [Key] = 1");
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
       public bool PETC0309IDelete(L_Log_Entity log_data)
       {
           return NyuukinMeisaiDL.PETC0303IDelete(log_data);

       }
       public bool CheckExistName1_0309I(T_NyuukinMeisai_Entity NyuukinMeisai_data, L_Log_Entity log_data, string twobytes)
       {
           if (NyuukinMeisaiDL.CheckExistName1_0309I(NyuukinMeisai_data, log_data, twobytes).Rows.Count > 0)
           {
               return true;
           }
           return false;
       }
       public bool PETC0309I_Insert(T_NyuukinMeisai_Entity NyuukinMeisai_data, L_Log_Entity log_data, string twobytes)
       {
           if (NyuukinMeisaiDL.PETC0309I_Insert(NyuukinMeisai_data, log_data, twobytes))
           {

               return true;

           }
           return false;
       }
    }
}
