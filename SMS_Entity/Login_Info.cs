using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SMS_Entity
{
    public class Login_Info
    {
        public string FormName { get; set; }
        public string PcLoginName { get; set; }
        public string PcName { get; set; }
        public string OperatorCode { get; set; }
        public string OperatorName { get; set; }
        public string CompanyCode { get; set; }
        public string CompanyName { get; set; }
        public string DepCode { get; set; }//department code
        public string DepName { get; set; }
        public string StoreCode { get; set; }
        public string StoreName { get; set; }
        public string ProcessDate { get; set; }//processing date
        public string CallerID { get; set; }
        public string StartMode { get; set; }
        public string Authority { get; set; }
        public string ItemCode { get; set; }
        public string ItemName { get; set; }
        public string type { get; set; }
        public string ProcessClass { get; set; }
    }
}
