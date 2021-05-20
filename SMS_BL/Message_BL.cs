using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SMS_Entity;
using SMS_DL;
using System.Data;

namespace SMS_BL
{
    public class Message_BL
    {
        SMSG_DL smsg_dl;
        public Message_BL()
        {
            smsg_dl = new SMSG_DL();
        }
        public bool INSERT_ErrorTable(string error,string formName)
        {
            return smsg_dl.INSERT_ErrorTable(error,formName );
        }

        public bool SMSG_Select(SMSG_Entity smsg_data,out string message1,out string message2,out string message3)
        {
            DataTable dt = smsg_dl.SMSG_Select(smsg_data);
            bool returnValue = false;

            if (dt.Rows.Count > 0)
            {
                message1 = dt.Rows[0]["MSG01"].ToString();
                message2 = dt.Rows[0]["MSG02"].ToString();
                message3 = dt.Rows[0]["MSG03"].ToString();
                returnValue = true;
            }
            else
            {
                message1 = message2 = message3 = string.Empty;
            }

            return returnValue;
        }
    }
}
