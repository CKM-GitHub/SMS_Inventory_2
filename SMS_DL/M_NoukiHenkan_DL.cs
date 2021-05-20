using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using SMS_Entity;

namespace SMS_DL
{
    public class M_NoukiHenkan_DL : Base_DL 
    {
        public bool M_NoukiHenkan_Insert(M_NoukiHenkan_Entity nouki_data)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            dic.Add("@PatternCD", nouki_data.PatternCD);
            dic.Add("@NoukiXML", nouki_data.NoukiXML);
            dic.Add("@InsertOperator", nouki_data.InsertOperator);
            dic.Add("@InsertDateTime", nouki_data.InsertDateTime);
            dic.Add("@UpdateOperator", nouki_data.UpdateOperator);
            dic.Add("@UpdateDateTime", nouki_data.UpdateDateTime);

            return InsertUpdateDeleteData(dic, "M_NoukiHenkan_Insert");
        }

        public bool M_NoukiHenkan_Delete(M_NoukiHenkan_Entity nouki_data)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@PatternCD", nouki_data.PatternCD);
            return InsertUpdateDeleteData(dic, "M_NoukiHenkan_Delete");
        }

        public DataTable M_NoukiHenkan_Select(M_NoukiHenkan_Entity nouki_data)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@PatternCD", nouki_data.PatternCD);
            return SelectData(dic, "M_NoukiHenkan_Select");
        }
    }
}
