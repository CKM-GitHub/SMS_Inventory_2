using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SMS_DL;
using SMS_Entity;

namespace SMS_BL
{
   public class PSKS0119M_BL:Base_BL
    {
       M_Zaiko_Kanri_DL mzkdl;
       public PSKS0119M_BL()
       {
           mzkdl = new M_Zaiko_Kanri_DL();
       }

       public DataTable SelectData(DataTable dtImport,string option)
       {
           string xmlImport = DataTableToXml(dtImport);
           return mzkdl.SelectData(xmlImport,option);
       }
       public DataTable SelectDataByAdminCD(string adminCD)
       {
           return mzkdl.SelectDataByAdminCD(adminCD);
       }
       public bool M_Zaiko_Kanri_Setting(M_Zaiko_Kanri_Entity mzke)
       {
           return mzkdl.M_Zaiko_Kanri_Setting(mzke);
       }

       public bool M_Zaiko_KanriInsertUpdate(M_Zaiko_Kanri_Entity mzke,L_Log_Entity lle)
       {
           mzke.xml = DataTableToXml(mzke.dt);
           return mzkdl.M_Zaiko_KanriInsertUpdate(mzke, lle);
       }

       public bool M_Zaiko_RenkanriDelete(M_Zaiko_Kanri_Entity mzke, L_Log_Entity lle)
       {
           return mzkdl.M_Zaiko_RenkanriDelete(mzke,lle);
       }

       public DataTable SelectDataByMakerCD(string taisho, string vc_shiiresaki, string vc_brand,string operatorCD)
       {
           return mzkdl.SelectDataByMakerCD(taisho, vc_shiiresaki, vc_brand, operatorCD);
       }

       public bool CheckMakerJanCD(DataTable dtImport)
       {
           string xmlImport = DataTableToXml(dtImport);
           DataTable dt=mzkdl.SelectData_ExistMakerJanCD(xmlImport);
           if (dt.Rows.Count > 0)
               return true;
           else
               return false;
       }
    }
}
