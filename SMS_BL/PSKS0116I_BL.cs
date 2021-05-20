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
    public class PSKS0116I_BL
    {
        M_MultiPorpose_DL mmpdl;
        Base_DL bdl;

        public PSKS0116I_BL()
        {
            mmpdl = new M_MultiPorpose_DL();
            bdl = new Base_DL();
        }

        public DataTable M_SiiresakiZaiko_Select(M_SiiresakiZaiko_Entity MSiireiSaki_Data)
        {
            M_SiiresakiZaiko_DL mszdl = new M_SiiresakiZaiko_DL();
            return mszdl.PSKS0116I_M_SiiresakiZaiko_Select(MSiireiSaki_Data);
        }

        public DataTable M_MultiPorpose_DynamicSelect(M_MultiPorpose_Entity mme)
        {
            DataTable dt = mmpdl.DynamicSelectData(mme.fields, mme.tableName, mme.condition);
            return dt;
        }

        public bool MakerBrandSelect(string fields, string tableName, string condition)
        {
            DataTable dt = new DataTable();
            dt = bdl.DynamicSelectData(fields, tableName, condition);
            if (dt.Rows.Count > 0)
                return true;
            else return false;
        }
    }
}
