using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using SMS_Entity;

namespace SMS_DL
{
   public class SuuryoHenKan_DL:Base_DL 
    {
        public DataTable SuuryoHenKan_Select(SuuryoHenKan_Entity suuryo_data)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@PatternCD", suuryo_data.PatternCD);
            return SelectData(dic, "SuuryoHenKan_Select");
        }

        public bool SuuryoHenKan_Insert(SuuryoHenKan_Entity suuryo_data)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            dic.Add("@PatternCD", suuryo_data.PatternCD);
            dic.Add("@SuuryoXML", suuryo_data.SuuryoXML);
            dic.Add("@InsertOperator", suuryo_data.InsertOperator);
            dic.Add("@InsertDateTime", suuryo_data.InsertDateTime);
            dic.Add("@UpdateOperator", suuryo_data.InsertOperator);
            dic.Add("@UpdateDateTime", suuryo_data.UpdateDateTime);

            return InsertUpdateDeleteData(dic, "SuuryoHenKan_Insert");
        }

        public bool SuuryoHenKan_Delete(SuuryoHenKan_Entity suuryo_data)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@PatternCD", suuryo_data.PatternCD);
            return InsertUpdateDeleteData(dic, "SuuryoHenKan_Delete");
        }
    }
}
