using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using SMS_Entity;
using System.Data.SqlClient;


namespace SMS_DL
{
    public class M_MakerZaiko_D_DL : Base_DL 
    {
        public bool M_MakerZaiko_D_Insert(MakerZaiko_D_Entity maker_D_entity)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            dic.Add("@PatternCD", maker_D_entity.PatternCD);
            dic.Add("@KamokuXML", maker_D_entity.KamokuXML);
            return InsertUpdateDeleteData(dic, "M_MakerZaiko_D_Insert");
        }

        public bool M_MakerZaiko_D_Delete(MakerZaiko_D_Entity maker_D_entity)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@PatternCD", maker_D_entity.PatternCD);
            return InsertUpdateDeleteData(dic, "M_MakerZaiko_D_Delete");
        }

        public DataTable M_MakerZaiko_D_Select(MakerZaiko_D_Entity maker_D_entity)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            dic.Add("@PatternCD", maker_D_entity.PatternCD);
            return SelectData(dic, "M_MakerZaiko_Select");
        }
    }
}
