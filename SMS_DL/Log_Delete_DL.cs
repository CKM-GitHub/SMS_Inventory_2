using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SMS_DL
{
    public class Log_Delete_DL:Base_DL
    {
        public bool Log_Delete()
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            return InsertUpdateDeleteData(dic, "PSKS0117B_Delete");
        }

        public bool SKSChangeFLG()
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            return InsertUpdateDeleteData(dic, "PSKS0118B_ChangeFLG");
        }
    }
}
