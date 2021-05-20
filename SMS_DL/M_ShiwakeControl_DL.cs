using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using SMS_Entity;

namespace SMS_DL
{
    public class M_ShiwakeControl_DL : Base_DL
    {

        public DataTable GetTaxRate(string date)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            dic.Add("@docDate", date);

            return SelectData(dic, "Get_TaxRate");
        }
        public DataTable GetTaxRate_MF(string date)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            dic.Add("@docDate", date);

            return SelectData(dic, "Get_TaxRate_MF");
        }
        //GetTaxRate_MF
    }
}
