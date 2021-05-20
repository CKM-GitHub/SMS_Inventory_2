using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SMS_Entity;
using SMS_DL;

namespace SMS_BL
{
    public class Menu_BL
    {
        Menu_DL menu_dl;
        public Menu_BL()
        {
            menu_dl = new Menu_DL();
        }

        public DataTable Menu_Select(Login_Info logininfo)
        {
            return menu_dl.Menu_Select(logininfo);
        }
    }
}
