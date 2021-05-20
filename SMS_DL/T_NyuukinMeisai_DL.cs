using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using SMS_Entity;

namespace SMS_DL
{
    public class T_NyuukinMeisai_DL : Base_DL
    {

        // public DataTable   PETC0302I_CSVExport1 
        public DataTable CheckExistName1_0309I(T_NyuukinMeisai_Entity NyuukinMeisai_data, L_Log_Entity lle, string twobytes)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            dic.Add("@OperatorCD", lle.InsertOperator);
            dic.Add("@date", lle.InsertDateTime);
            dic.Add("@Filename", twobytes);
            dic.Add("@XML", NyuukinMeisai_data.xml);

            return SelectData(dic, "CheckExistName1_0309I");

        }
        public DataTable CheckExistName1_0303I(T_NyuukinMeisai_Entity NyuukinMeisai_data, L_Log_Entity lle, string twobytes)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            dic.Add("@OperatorCD", lle.InsertOperator);
            dic.Add("@date", lle.InsertDateTime);
            dic.Add("@Filename", twobytes);
            dic.Add("@XML", NyuukinMeisai_data.xml);

            return SelectData(dic, "CheckExistName1_0303I");

        }

        public bool PETC0303IDelete(L_Log_Entity lle)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            dic.Add("@OperatorCD", lle.InsertOperator);
            return InsertUpdateDeleteData(dic, "PETC0303I_Delete_Loop");
          
        }
        public bool PETC0309IDelete(L_Log_Entity lle)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            dic.Add("@OperatorCD", lle.InsertOperator);
            return InsertUpdateDeleteData(dic, "PETC0309I_Delete_Loop");

        }

        public bool PETC0308I_Insert(T_NyuukinMeisai_Entity NyuukinMeisai_data, L_Log_Entity lle, string twobytes)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            dic.Add("@OperatorCD", lle.InsertOperator);
            dic.Add("@date", lle.InsertDateTime);
            dic.Add("@Filename", twobytes);
            dic.Add("@XML", NyuukinMeisai_data.xml);

            return InsertUpdateDeleteData(dic, "PETC0303I_Insert");

        }
        public bool PETC0303I_Insert(T_NyuukinMeisai_Entity NyuukinMeisai_data, L_Log_Entity lle,string twobytes)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            dic.Add("@OperatorCD", lle.InsertOperator);
            dic.Add("@date", lle.InsertDateTime);
            dic.Add("@Filename", twobytes);
            dic.Add("@XML", NyuukinMeisai_data.xml);

            return InsertUpdateDeleteData(dic, "PETC0303I_Insert");
          
        }
        public bool PETC0309I_Insert(T_NyuukinMeisai_Entity NyuukinMeisai_data, L_Log_Entity lle, string twobytes)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            dic.Add("@OperatorCD", lle.InsertOperator);
            dic.Add("@date", lle.InsertDateTime);
            dic.Add("@Filename", twobytes);
            dic.Add("@XML", NyuukinMeisai_data.xml);

            return InsertUpdateDeleteData(dic, "PETC0309I_Insert");

        }
        public bool T_NyuukinMeisai_Delete(T_NyuukinMeisai_Entity NyuukinMeisai_data)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            dic.Add("@OperatorCD", NyuukinMeisai_data.OperatorCD);

            return InsertUpdateDeleteData(dic, "T_NyuukinMeisai_Delete");
        }

        public bool T_NyuukinMeisai_Insert(T_NyuukinMeisai_Entity NyuukinMeisai_data)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            dic.Add("@OperatorCD", NyuukinMeisai_data.OperatorCD);
            dic.Add("@date", NyuukinMeisai_data.InsertDatetime);
            dic.Add("@XML", NyuukinMeisai_data.xml);

            return InsertUpdateDeleteData(dic, "T_NyuukinMeisai_Insert");
        }
        public bool T_NyuukinMeisai_Insert_MF(T_NyuukinMeisai_Entity NyuukinMeisai_data)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            dic.Add("@OperatorCD", NyuukinMeisai_data.OperatorCD);
            dic.Add("@date", NyuukinMeisai_data.InsertDatetime);
            dic.Add("@XML", NyuukinMeisai_data.xml);

            return InsertUpdateDeleteData(dic, "T_NyuukinMeisai_Insert_MF");
        }
        public DataTable PETC0302I_Detail_SELECT(T_NyuukinMeisai_Entity NyuukinMeisai_data)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            dic.Add("@OperatorCD", NyuukinMeisai_data.OperatorCD);

            return SelectData(dic, "PETC0302I_Detail_SELECT");
        }
        public DataTable PETC0308I_Detail_SELECT(T_NyuukinMeisai_Entity NyuukinMeisai_data)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            dic.Add("@OperatorCD", NyuukinMeisai_data.OperatorCD);

            return SelectData(dic, "PETC0308I_Detail_SELECT");
        }
        public DataTable PETC0302I_CSVExport1(T_NyuukinMeisai_Entity NyuukinMeisai_data, string expname)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            dic.Add("@OperatorCD", NyuukinMeisai_data.OperatorCD);
            return SelectData(dic, expname);
        }

        public DataTable Check_Motonames(T_NyuukinMeisai_Entity NyuukinMeisai_data)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            dic.Add("@OperatorCD", NyuukinMeisai_data.OperatorCD);
            dic.Add("@xml", NyuukinMeisai_data.xml);

            return SelectData(dic, "Check_Motonames");
        }
        public DataTable Check_Motonames_MF(T_NyuukinMeisai_Entity NyuukinMeisai_data)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            dic.Add("@OperatorCD", NyuukinMeisai_data.OperatorCD);
            dic.Add("@xml", NyuukinMeisai_data.xml);

            return SelectData(dic, "Check_Motonames_MF");
        }
        public bool T_NyuukinMeisai_Tesuuryou_Delete(L_Log_Entity lle)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            dic.Add("@OperatorCD", lle.InsertOperator);
            return InsertUpdateDeleteData(dic, "PETC0308I_Delete_Loop");

        }
        //T_NyuukinMeisai_Tesuuryou_Delete
    }
}
