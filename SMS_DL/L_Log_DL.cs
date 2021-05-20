using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SMS_Entity;

namespace SMS_DL
{
    public class L_Log_DL : Base_DL
    {
        public bool L_Log_Insert(L_Log_Entity L_LogEntity)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@OperateDate", L_LogEntity.OperateDate);
            dic.Add("@InsertOperator", L_LogEntity.InsertOperator);
            dic.Add("@Program", L_LogEntity.Program);
            dic.Add("@PC", L_LogEntity.PC);
            dic.Add("@OperateMode", L_LogEntity.OperateMode);
            dic.Add("@KeyItem", L_LogEntity.KeyItem);
            return InsertUpdateDeleteData(dic, "L_Log_Insert");
        }
    }
}
