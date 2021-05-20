using SMS_Entity;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SMS_DL
{
    public class T_Zaiko_DL : Base_DL
    {
        public bool T_Zaiko_INSERT_UPDATE(T_Zaiko_Entity Zaiko_Data)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@HanbaiShohinCD", Zaiko_Data.HanbaiShohinCD);
            dic.Add("@JanCD", Zaiko_Data.JanCD);
            dic.Add("@CapitalZaiko", Zaiko_Data.CapitalZaiko.Replace(",",""));
            dic.Add("@ToyonakaZaiko", Zaiko_Data.ToyonakaZaiko.Replace(",", ""));
            dic.Add("@IshibashiZaiko", Zaiko_Data.IshibashiZaiko.Replace(",", ""));
            dic.Add("@EsakaZaiko", Zaiko_Data.EsakaZaiko.ToString().Replace(",", ""));
            dic.Add("@SannomiyaZaiko", Zaiko_Data.SannomiyaZaiko.ToString().Replace(",", ""));
            dic.Add("@UpdateDateTime", Zaiko_Data.UpdateDateTime.ToString());

            return InsertUpdateDeleteData(dic, "T_Zaiko_INSERT_UPDATE");
        }

        public bool PSKS0115B1_TZaiko_INSERT_UPDATE()
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            //return InsertUpdateDeleteData(dic, "PSKS0115B1_Select1");
            return InsertUpdateDeleteData(dic, "PSKS0115B1_TZaiko_INSERT_UPDATE");//att update
        }

        public bool PSKS0115B2_TZaiko_INSERT_UPDATE()
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            return InsertUpdateDeleteData(dic, "PSKS0115B2_TZaiko_INSERT_UPDATE");//att update
        }
    }
}
