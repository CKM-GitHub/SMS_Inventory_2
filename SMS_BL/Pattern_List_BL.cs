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
    public class Pattern_List_BL : Base_BL //Pattern_List
    {
        M_MakerZaiko_H_DL maker_H_DL = new M_MakerZaiko_H_DL();

        public DataTable M_MakerZaiko_H_Select(M_MakerZaiko_H_Entity maker_H_data)
        {
            return maker_H_DL.M_MakerZaiko_H_Select(maker_H_data);
        }
    }
}
