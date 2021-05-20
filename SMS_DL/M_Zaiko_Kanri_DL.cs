using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SMS_Entity;

namespace SMS_DL
{
    public class M_Zaiko_Kanri_DL:Base_DL
    {
        public DataTable SelectData(string xmlImport,string option)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@xml", xmlImport);
            dic.Add("@option", option);
            return SelectData(dic, "PSKS0119M_SelectData");
        }

        public DataTable SelectData_ExistMakerJanCD(string xmlImport)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@xml", xmlImport);
            return SelectData(dic, "PSKS0119M_CheckMakerJanCD");
        }

        public DataTable SelectDataByAdminCD(string adminCD)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@AdminCD", adminCD);
            return SelectData(dic, "PSKS0119M_SelectDataByAdminCD");
        }

        public bool M_Zaiko_Kanri_Setting(M_Zaiko_Kanri_Entity mzke)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@OperatorCD", mzke.OperatorCD);
            dic.Add("@AdminCD", mzke.AdminCD);
            dic.Add("@taisho", mzke.taisho);
            dic.Add("@kijunsu", mzke.kijunsu);
            dic.Add("@Zaiko_keisan", mzke.Zaiko_keisan);
            return InsertUpdateDeleteData(dic, "PSKS0119M_InsertUpdate");

           
        }

        public DataTable PSKS0120S_InitialSelect(M_Zaiko_Kanri_Entity zaiko_kanri)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            dic.Add("@OperatorCD", zaiko_kanri.OperatorCD);

            return SelectData(dic, "PSKS0120S_InitialSelect");
        }

        public bool PSKS0120S_Update(M_Zaiko_Kanri_Entity zaiko_kanri)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            dic.Add("@xml", zaiko_kanri.xml);
            dic.Add("@UpdateDateTime",zaiko_kanri.UpdateDateTime);

            return InsertUpdateDeleteData(dic, "PSKS0120S_Update");
        }

        public bool M_Zaiko_KanriInsertUpdate(M_Zaiko_Kanri_Entity mzke,L_Log_Entity lle)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@xml", mzke.xml);
            dic.Add("@Program", lle.Program);
            dic.Add("@PC", lle.PC);
            dic.Add("@OperateMode", lle.OperateMode);
            return InsertUpdateDeleteData(dic, "M_Zaiko_KanriInsertUpdate");
        }

        public bool M_Zaiko_RenkanriDelete(M_Zaiko_Kanri_Entity mzke, L_Log_Entity lle)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@OperatorCD", mzke.OperatorCD);
            dic.Add("@AdminCD", mzke.AdminCD);
            dic.Add("@Program", lle.Program);
            dic.Add("@PC", lle.PC);
            dic.Add("@OperateMode", lle.OperateMode);
            return InsertUpdateDeleteData(dic, "M_Zaiko_RenkanriDelete");
        }


        public DataTable SelectDataByMakerCD(string taisho, string vc_shiiresaki, string vc_brand, string operatorCD)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@taisho", taisho);
            dic.Add("@vc_shiiresaki",vc_shiiresaki);
            dic.Add("@vc_brand", vc_brand);
            dic.Add("@OperatorCD", operatorCD);
            return SelectData(dic, "PSKS0119M_SelectDataByMakerCD");
        }
    }
}
