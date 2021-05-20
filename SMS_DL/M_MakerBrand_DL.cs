using SMS_Entity;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SMS_DL
{
    public class M_MakerBrand_DL : Base_DL
    {
        public DataTable M_MakerBrand_Select(M_MakerBrand_Entity mmkbData)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@MakerCD", mmkbData.MakerCD);
            dic.Add("@BrandCD", mmkbData.BrandCD);
            dic.Add("@DatasourceMakerCD", mmkbData.DatasourceMakerCD);
            return SelectData(dic, "M_MakerBrand_Select");
        }

        public bool M_MakerBrand_InsertXML(M_MakerBrand_Entity mmbe)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@MakerCD", mmbe.MakerCD);
            dic.Add("@BrandCD", mmbe.BrandCD);
            dic.Add("@DatasourceMakerCD",mmbe.DatasourceMakerCD);
            dic.Add("@InsertOperator", mmbe.InsertOperator);
            dic.Add("@InsertDateTime", mmbe.InsertDateTime);
            dic.Add("@XML", mmbe.xml);
            return InsertUpdateDeleteData(dic, "M_MakerBrand_InsertXML");
        }
    }
}
