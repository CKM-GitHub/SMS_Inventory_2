using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using SMS_Entity;

namespace SMS_DL
{
    public class M_MakerZaiko_H_DL : Base_DL 
    {
        public bool M_MakerZaiko_H_Update(M_MakerZaiko_H_Entity maker_H_data)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            dic.Add("@PatternCD", maker_H_data.PatternCD);
            dic.Add("@PatternName", maker_H_data.PatternName);
            dic.Add("@InsertOperator", maker_H_data.InsertOperator);
            dic.Add("@InsertDateTime", maker_H_data.InsertDateTime);
            dic.Add("@UpdateOperator", maker_H_data.InsertOperator);
            dic.Add("@UpdateDateTime", maker_H_data.InsertDateTime);


            return InsertUpdateDeleteData(dic, "M_MakerZaiko_H_Update");
        }

        public bool M_MakerZaiko_H_Delete(M_MakerZaiko_H_Entity maker_H_data)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            dic.Add("@PatternCD", maker_H_data.PatternCD);
            return InsertUpdateDeleteData(dic, "MakerZaiko_H_Delete");
        }

        public DataTable M_MakerZaiko_H_Select(M_MakerZaiko_H_Entity maker_H_data) //Pattern_List_BL
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            dic.Add("@PatternCD", maker_H_data.PatternCD);
            dic.Add("@PatternName", maker_H_data.PatternName);

            return SelectData(dic, "M_MakerZaiko_H_Select");
        }
    }
}
