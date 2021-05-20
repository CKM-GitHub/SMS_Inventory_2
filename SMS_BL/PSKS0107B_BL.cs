using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using SMS_Entity;
using SMS_DL;

namespace SMS_BL
{
    public class PSKS0107B_BL : Base_BL
    {
        Base_DL BaseDL = new Base_DL();
        T_Hanbai_Shohin_DL thanbaishohin_dl = new T_Hanbai_Shohin_DL();

        public DataTable M_MultiPorpose_SelectID(M_MultiPorpose_Entity MmultiporposeData)
        {
            return BaseDL.DynamicSelectData("Char1 ,Char2,Num1,Num2", "M_MultiPorpose", "where ID= " + MmultiporposeData.ID + " AND [Key]='" + MmultiporposeData.Key + "'");
        }

        public bool Data_Insert(Base_Entity baseEntity)
        {
            return thanbaishohin_dl.DataInsert(baseEntity);
        }
    }
}
