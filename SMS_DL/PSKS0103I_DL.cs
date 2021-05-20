using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SMS_Entity;
using System.Data;

namespace SMS_DL
{
    public class PSKS0103I_DL : Base_DL
    {
        public bool Import(T_MakerInport_Entity tmie, string insertDateTime,string noukiFlag)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@DataSourseMakerCD", tmie.DataSourseMakerCD);
            dic.Add("@InportFile", tmie.InportFile);
            dic.Add("@PatternCD", tmie.PatternCD);
            dic.Add("@InsertOperator", tmie.InsertOperator);
            dic.Add("@InsertDateTime", insertDateTime);
            dic.Add("@NoukiFlg", noukiFlag);
            dic.Add("@XML", tmie.xml);
            return InsertUpdateDeleteData(dic, "PSKS0103I_Import");
        }

        public bool ImportProcess(T_MakerInport_Entity tmie, string insertDateTime)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@DataSourseMakerCD", tmie.DataSourseMakerCD);
            dic.Add("@InportFile", tmie.InportFile);

            dic.Add("@PatternCD", tmie.PatternCD);
            dic.Add("@InsertOperator", tmie.InsertOperator);
            dic.Add("@InsertDateTime", insertDateTime);
            dic.Add("@XML", tmie.xml);
            return InsertUpdateDeleteData(dic, "PSKS01013I_ImportProcess");
        } 

        public bool PSKS0103I_DELETE(T_MakerInport_Entity tmie)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@InportSEQ", tmie.InportSEQ);
            dic.Add("@InsertOperator", tmie.InsertOperator);
            dic.Add("@InsertDateTime", tmie.InsertDateTime);

            return InsertUpdateDeleteData(dic, "PSKS0103I_DELETE");
        }

        public string JANCD_PositionCheck(string MakerCD)
        {
            DataTable dt = SelectData_CommandText("select mzd.SEQ from M_MakerZaiko_D MZD " + " inner join M_MakerBrand MB " + " on mzd.PatternCD = MB.PatternCD " + " where MB.MakerCD = '" + MakerCD + "' " + " and mzd.ItemProperty = 1");

            if (dt.Rows.Count > 0)
                return dt.Rows[0]["SEQ"].ToString();
            else
                return "0";
        }

        public DataTable T_MakerZaiko_ZaikoSuTotal(string oldMakerCD)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@DataMoto", oldMakerCD);
            return SelectData(dic, "T_MakerZaiko_ZaikoSuTotal");
        }

        public DataTable Nouki_Select(string Nouki)
        {
            return SelectData_CommandText("select HenkanNouki from M_NoukiHenkan MNH inner join M_MakerBrand MB "
                                        + " on MNH.PatternCD = MB.PatternCD and MNH.Nouki = '" + Nouki + "'");
        }
    }
}
