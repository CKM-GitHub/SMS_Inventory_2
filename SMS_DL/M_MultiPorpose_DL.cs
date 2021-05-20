using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SMS_Entity;
using System.Data;
using System.Collections;
using System.Data.SqlClient;

namespace SMS_DL
{
    public class M_MultiPorpose_DL : Base_DL
    {
        public DataTable get_ptnname(string mkr_CD)
        {

            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@mkr_CD", mkr_CD);
         //   dic.Add("@vm_bunrui", mbe.vm_bunrui);
            return SelectData(dic, "Get_Ptn_CD");
        }

        public bool M_MultiPorpose_Num1_Update(M_MultiPorpose_Entity MmultiporposeData)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@ID", MmultiporposeData.ID);
            dic.Add("@Num1", MmultiporposeData.Num1);
            dic.Add("@UpdateOperator", MmultiporposeData.InsertOperator);
            dic.Add("@UpdateDateTime", MmultiporposeData.InsertDateTime);

            return InsertUpdateDeleteData(dic, "M_MultiPorpose_Num1_Update");
        }
    }
}
