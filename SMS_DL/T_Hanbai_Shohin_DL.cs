using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SMS_Entity;

namespace SMS_DL
{
    public class T_Hanbai_Shohin_DL : Base_DL
    {
        public bool DataInsert(Base_Entity baseEntity)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@DateTime", baseEntity.InsertDateTime);
            return InsertUpdateDeleteData(dic, "PSKS0107B_Insert");
        }
    }
}
