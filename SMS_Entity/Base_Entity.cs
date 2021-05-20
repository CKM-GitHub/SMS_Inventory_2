using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace SMS_Entity
{
    public class Base_Entity
    {
        public string KAICD { get; set; }
        public string InsertOperator { get; set; }
        public string InsertDateTime { get; set; }
        public string xml { get; set; }
        public string fields { get; set; }
        public string tableName { get; set; }
        public string condition { get; set; }
        
        public DataTable dt1 { get; set; }
    }
}
