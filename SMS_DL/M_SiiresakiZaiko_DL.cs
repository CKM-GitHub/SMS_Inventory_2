using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SMS_Entity;

namespace SMS_DL
{
    public class M_SiiresakiZaiko_DL : Base_DL
    {
        public DataTable M_SiiresakiZaiko_Select(M_SiiresakiZaiko_Entity msze)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@SiiresakiCD", msze.SiiresakiCD);
            return SelectData(dic, "M_SiiresakiZaiko_Select");
        }

        public DataTable PSKS0116I_M_SiiresakiZaiko_Select(M_SiiresakiZaiko_Entity msze)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@SiiresakiCD", msze.SiiresakiCD);
            return SelectData(dic, "PSKS0116I_M_SiiresakiZaiko_Select");
        }

        public bool M_SiiresakiZaiko_Update(M_SiiresakiZaiko_Entity msze)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@MakerCD", msze.SiiresakiCD);
            dic.Add("@KakuTeiFlg", msze.KakuTeiFlg);
            dic.Add("@InsertDateTime", msze.InsertDateTime);
            dic.Add("@FileKBN", msze.FileKBN);
            return InsertUpdateDeleteData(dic, "M_SiiresakiZaiko_Update");
        }
    }
}
