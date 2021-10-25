using SMS_Entity;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SMS_DL {
    public class Excel_DL : Base_DL{
        public DataTable Excel_Select(string val1,string val2)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            dic.Add("@start", val1);
            dic.Add("@end", val2);

            return SelectData(dic, "Excel_Select");
        }
        public DataTable Mail_Select()
        {
            string sp = "Mail_Select";
            Dictionary<string, string> dic = new Dictionary<string, string>();
           
            return SelectData(dic, sp);
        }
      
        public bool MailSend_Update(int MailCount)
        {
            string sp = "Mail_Send_Update";
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@MailCount", MailCount.ToString());   
            
            return InsertUpdateDeleteData(dic, sp);
        }
    }
}
