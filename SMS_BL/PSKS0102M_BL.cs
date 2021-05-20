using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SMS_DL;
using SMS_Entity;
using System.Data;

namespace SMS_BL
{
    public class PSKS0102M_BL : Base_BL
    {
        M_MakerBrand_DL makerbrandDL;
        M_SHIIRESAKI_DL shiiresakiDL;
        M_Brand_DL brandDL;
        M_MakerZaiko_H_DL makerzaikohDL;
        L_Log_DL LogDL;
        public PSKS0102M_BL()
        {
            makerbrandDL = new M_MakerBrand_DL();
            shiiresakiDL = new M_SHIIRESAKI_DL();
            brandDL = new M_Brand_DL();
            makerzaikohDL = new M_MakerZaiko_H_DL();
            LogDL = new L_Log_DL();
        }

        public DataTable M_MakerBrand_Select(M_MakerBrand_Entity mmkbData)
        {
            return makerbrandDL.M_MakerBrand_Select(mmkbData);
        }

        public DataTable M_Shiiresaki_Select(M_SHIIRESAKI_Entity mse)
        {
            return shiiresakiDL.M_Shiiresaki_Select(mse);
        }

        public DataTable M_Brand_Select(M_Brand_Entity mbe)
        {
            return brandDL.M_Brand_Select(mbe);
        }

        public DataTable M_MakerZaiko_H_Select(M_MakerZaiko_H_Entity mhe) //Pattern_List_BL
        {
            return makerzaikohDL.M_MakerZaiko_H_Select(mhe);
        }

        public bool M_MakerBrand_Insert_Xml(M_MakerBrand_Entity mmbe,L_Log_Entity lle)
        {
            mmbe.xml = DataTableToXml(mmbe.dt1);

            makerbrandDL.StartTransaction();

            if (makerbrandDL.M_MakerBrand_InsertXML(mmbe))
            {
                LogDL.Transaction = makerbrandDL.Transaction;
                if (LogDL.L_Log_Insert(lle))
                {
                    makerbrandDL.CommitTransaction();
                    return true;
                }
            }

            return false;
        }
    }
}
