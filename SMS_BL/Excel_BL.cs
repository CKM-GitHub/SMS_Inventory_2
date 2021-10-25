using SMS_DL;
using SMS_Entity;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace SMS_BL {
    public class Excel_BL :Base_BL{
        Excel_DL excel_DL = new Excel_DL();

        public DataTable Excel_Select(string val1,string val2)
        {
            return excel_DL.Excel_Select(val1,val2);
        }
        public DataTable Mail_Select()
        {
            return excel_DL.Mail_Select();
        }
        public bool MailSend_Update(int MailCount)
        {
            return excel_DL.MailSend_Update(MailCount);
        }
    }
}
