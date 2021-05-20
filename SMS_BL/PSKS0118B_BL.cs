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
    public class PSKS0118B_BL
    {
        Base_DL BaseDL = new Base_DL();
        Log_Delete_DL log_del = new Log_Delete_DL();

        public DataTable M_MultiPorpose_Select(string id)
        {
            return BaseDL.DynamicSelectData("Num1,Num2", "M_MultiPorpose", "where ID=" + id + " AND [Key]='1'");
        }


        public bool SKSChangeFLG()
        {
            return log_del.SKSChangeFLG();
        }
    }
}
