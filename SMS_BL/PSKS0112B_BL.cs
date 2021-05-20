using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Threading.Tasks;
using SMS_Entity;
using System.Data.SqlClient;
using SMS_DL;

namespace SMS_BL
{
    public class PSKS0112B_BL : Base_BL
    {
        M_MultiPorpose_DL MmultiporposeDL = new M_MultiPorpose_DL();

        PSKS0112B_DL psks0112bdl = new PSKS0112B_DL();
        
        public DataTable M_MultiPorpose_Select()
        {
            return MmultiporposeDL.DynamicSelectData("Num1,Num2", "M_MultiPorpose", "where ID=109 AND [Key]='1'");
        }

        public DataTable PSKS0112B_Select(Base_Entity be)
        {
            return psks0112bdl.PSKS0112B_Select(be);
        }
    }
}
