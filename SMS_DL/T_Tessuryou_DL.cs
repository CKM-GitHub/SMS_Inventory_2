using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SMS_Entity;
using System.Data;

namespace SMS_DL
{
    public class T_Tessuryou_DL : Base_DL
    {
        public bool Update_Tesuyou_Nyuukin(string tesuyou, string nyuukin)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            dic.Add("@xmltesu", tesuyou);
            dic.Add("@xmlnyuukin", nyuukin);
            return InsertUpdateDeleteData(dic, "Update_Tesuyou_Nyuukin");
        
        }
        public bool Update_Tesuyou_Nyuukin_MF_309I(string tesuyou, string nyuukin)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            dic.Add("@xmltesu", tesuyou);
            dic.Add("@xmlnyuukin", nyuukin);
            return InsertUpdateDeleteData(dic, "Update_Tesuyou_Nyuukin_MF_309I");

        }
        public bool Update_Tesuyou_Nyuukin_MF(string tesuyou, string nyuukin)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            dic.Add("@xmltesu", tesuyou);
            dic.Add("@xmlnyuukin", nyuukin);
            return InsertUpdateDeleteData(dic, "Update_Tesuyou_Nyuukin_MF");

        }
        //Update_Tesuyou_Nyuukin_MF
        public DataSet PETC0303I_SELECT(string opcd)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            dic.Add("@opcd", opcd);

            return SelectDataSet(dic, "PETC0303i_Select");

        }
        public DataSet PETC0309I_SELECT(string opcd)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            dic.Add("@opcd", opcd);

            return SelectDataSet(dic, "PETC0303i_Select");

        }
        public DataTable PETC0302I_SELECT(T_Tesuuryou_Entity tse)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            dic.Add("@opcd",tse.OperatorCD );

            return SelectData(dic, "PETC0302I_SELECT");
        }
        public DataTable PETC0308I_SELECT_MF(T_Tesuuryou_Entity tse)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            dic.Add("@opcd", tse.OperatorCD);

            return SelectData(dic, "PETC0308I_SELECT");
        }
        public bool T_Tesuuryou_Delete(T_Tesuuryou_Entity Tesuuryou_data)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            dic.Add("@OperatorCD", Tesuuryou_data.OperatorCD);

            return InsertUpdateDeleteData(dic, "T_Tesuuryou_Delete");
        }

        public bool T_Tesuuryou_Insert(T_Tesuuryou_Entity Tesuuryou_data)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            dic.Add("@OperatorCD", Tesuuryou_data.OperatorCD);
            dic.Add("@date", Tesuuryou_data.InsertDateTime);

            return InsertUpdateDeleteData(dic, "T_Tesuuryou_Insert");
        }
        public bool T_Tesuuryou_Insert_MF(T_Tesuuryou_Entity Tesuuryou_data)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            dic.Add("@OperatorCD", Tesuuryou_data.OperatorCD);
            dic.Add("@date", Tesuuryou_data.InsertDateTime);

            return InsertUpdateDeleteData(dic, "T_Tesuuryou_Insert_MF");
        }
    }
}
