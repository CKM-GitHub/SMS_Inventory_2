using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using SMS_DL;
using SMS_Entity;

namespace SMS_BL
{
    public class PSKS0115B_BL : Base_BL
    {
        M_MultiPorpose_DL m_multiporposedl = new M_MultiPorpose_DL();
        T_Zaiko_DL tz_dl = new T_Zaiko_DL();
        Base_DL BaseDL = new Base_DL();

        public DataTable M_MultiPorpose_Select(string id)
        {
            return BaseDL.DynamicSelectData("Num1,Num2", "M_MultiPorpose", "where ID="+id+" AND [Key]='1'");
        }

        public bool PSKS0115B1_TZaiko_INSERT_UPDATE()
        {
            return tz_dl.PSKS0115B1_TZaiko_INSERT_UPDATE();
        }

        public bool PSKS0115B2_TZaiko_INSERT_UPDATE()
        {
            return tz_dl.PSKS0115B2_TZaiko_INSERT_UPDATE();
        }
    }
}
