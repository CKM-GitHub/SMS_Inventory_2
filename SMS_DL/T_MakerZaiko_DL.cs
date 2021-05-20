using SMS_Entity;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SMS_DL
{
    public class T_MakerZaiko_DL : Base_DL
    {
        public bool T_MakerZaiko_DELETE(T_MakerZaiko_Entity tmze)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@MakerCD", tmze.MakerCD);

            return InsertUpdateDeleteData(dic, "T_MakerZaiko_DELETE");
        }

        public bool T_MakerZaiko_Insert(T_MakerZaiko_Entity tmze)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@DateTime", tmze.InsertDateTime);
            dic.Add("@MakerCD", tmze.MakerCD);
            dic.Add("@InsertOperator", tmze.InsertOperator);
            return InsertUpdateDeleteData(dic, "T_MakerZaiko_INSERT");
        }

        public bool T_MakerZaiko_Update(T_MakerZaiko_Entity tmze)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@MakerCD", tmze.MakerCD);
            dic.Add("@KakuTeiFlg", tmze.KakuTeiFlg);
            dic.Add("@xml", tmze.xml);
            dic.Add("@InsertOperator", tmze.InsertOperator);
            dic.Add("@InsertDateTime", tmze.InportDateTime);
             
            return InsertUpdateDeleteData(dic, "T_MakerZaiko_Update");
        }

    }
}
