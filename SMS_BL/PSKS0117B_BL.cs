using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using SMS_Entity;
using SMS_DL;

namespace SMS_BL
{
    public class PSKS0117B_BL
    {
        Base_DL BaseDL = new Base_DL();
        Log_Delete_DL log_del = new Log_Delete_DL();

        public bool Log_Delete()
        {
            return log_del.Log_Delete();
        }
    }
}
