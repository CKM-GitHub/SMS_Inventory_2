using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SMS_Entity
{
   public class T_AmazonYoteiDate_Entity : Base_Entity
    {
        public string OperatorCD { get; set; }
        public string SEQ { get; set; }
        public string amazon_order_id { get; set; }
        public string sku { get; set; }
        public string estimated_ship_date { get; set; }
    }
}
