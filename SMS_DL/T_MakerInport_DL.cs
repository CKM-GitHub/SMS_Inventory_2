using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SMS_DL
{
    public class T_MakerInport_DL : Base_DL
    {
        public DataTable T_MakerInport_Select(string InportSEQ, string ErrChkFlg)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@InportSEQ", InportSEQ);
            dic.Add("@ErrChkFlg", ErrChkFlg);
            return SelectData(dic, "T_MakerInport_Select");
        }
    }
}
