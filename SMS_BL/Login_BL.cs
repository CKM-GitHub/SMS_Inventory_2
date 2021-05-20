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
    public class Login_BL 
    {
        MKAI_DL mkaiDL = new MKAI_DL();
        MOPE_Entity mopeData = new MOPE_Entity();
        MOPE_DL mopeDL = new MOPE_DL();
        SMNU_DL smnuDl = new SMNU_DL();

        public DataTable MKAI_SELECT(MKAI_Entity mkaiData) //Login
        {
            return mkaiDL.MKAI_SELECT(mkaiData);
        }

        public DataTable MOPE_SELECT(MOPE_Entity mopeData) //login
        {
            return mopeDL.MOPE_SELECT(mopeData);
        }

        public DataTable SMNU_SELECT(MOPE_Entity mopeData) //login
        {
            return smnuDl.SMNU_SELECT(mopeData);
        }

        public bool MOPE_UPDATE(MOPE_Entity mopeData) //LoginPassword
        {
            return mopeDL.MOPE_UPDATE(mopeData);
        }
    }
}
