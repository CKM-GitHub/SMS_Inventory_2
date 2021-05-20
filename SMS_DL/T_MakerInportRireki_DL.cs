using SMS_Entity;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SMS_DL
{
    public class T_MakerInportRireki_DL : Base_DL
    {
        public DataTable T_MakerInportRireki_SELECT(T_MakerInportRireki_Entity inport_data)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@DataSourseMakerCD", inport_data.DataSourseMakerCD);
            return SelectData(dic, "T_MakerInportRireki_SELECT");
        }
    }
}
