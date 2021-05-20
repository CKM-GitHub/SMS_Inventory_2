using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SMS_DL;
using SMS_Entity;
using System.Data;

namespace SMS_BL
{
    public class Base_BL
    {
        Common_DL commondl = new Common_DL();
        public DataTable DynamicSelectData(string fields, string tableName, string condition)
        {
            return commondl.DynamicSelectData(fields, tableName, condition);
        }

        public bool  DynamicUpdateData(M_MultiPorpose_Entity mme)
        {
            return commondl.DynamicUpdateData(mme);
        }

        public String DataTableToXml(DataTable dt)
        {
            dt.TableName = "test";
            System.IO.StringWriter writer = new System.IO.StringWriter();
            dt.WriteXml(writer, XmlWriteMode.WriteSchema, false);
            string result = writer.ToString();
            return result;
        }
    }
}
