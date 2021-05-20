using SMS_DL;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SMS_Entity;
using System.Data;
namespace SMS_BL
{
    public class PETC0306I_BL : Base_BL
    {
        T_AmazonYoteiDate_DL AmazonYoteidateDL;
        T_AmazonYoteiDate_Entity amazonyoteidate;
        M_MultiPorpose_DL mdl;
        L_Log_DL LogDL;

        public PETC0306I_BL()
        {
            AmazonYoteidateDL = new T_AmazonYoteiDate_DL();
            mdl = new M_MultiPorpose_DL();
            LogDL = new L_Log_DL();
        }
        public bool PETC0306I_DELETE(T_AmazonYoteiDate_Entity tmie)
        {
            return AmazonYoteidateDL.T_AmazonYoteiDate_Delete(tmie);
        }

        public bool PETC0306I_Insert(T_AmazonYoteiDate_Entity tmie)
        {
            return AmazonYoteidateDL.T_AmazonYoteiDate_Insert(tmie);
        }
        public bool PETC0306I_Update(T_AmazonYoteiDate_Entity tmie)
        {
            return AmazonYoteidateDL.T_AmazonYoteiDate_Update(tmie);
        }
        public DataTable PETC0306I_Select(T_AmazonYoteiDate_Entity tmie)
        {
            return AmazonYoteidateDL.T_AmazonYoteiDate_Select(tmie);
        }
        
    }
}
