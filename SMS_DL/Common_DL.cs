using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using SMS_Entity;

namespace SMS_DL
{
    public class Common_DL : Base_DL
    {
        public DataTable DynamicSelect(string fields, string tableName, string condition)
        {
            return DynamicSelectData(fields, tableName, condition);
        }
    }
}
