using SMS_DL;
using SMS_Entity;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace SMS_BL
{
    public class T_AmazonYoteiDate_DL : Base_DL
    {
        public bool T_AmazonYoteiDate_Delete(T_AmazonYoteiDate_Entity amazonyoteidate)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@OperatorCD", amazonyoteidate.OperatorCD);
            return InsertUpdateDeleteData(dic, "T_AmazonYoteiDate_Delete");
        }

        public bool T_AmazonYoteiDate_Insert(T_AmazonYoteiDate_Entity amazonyoteidate)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            dic.Add("@OperatorCD", amazonyoteidate.OperatorCD);
            dic.Add("@SEQ", amazonyoteidate.SEQ);
            dic.Add("@amazon_order_id", amazonyoteidate.amazon_order_id);
            dic.Add("@sku", amazonyoteidate.sku);
            return InsertUpdateDeleteData(dic, "T_AmazonYoteiDate_Insert");
        }
        public bool T_AmazonYoteiDate_Update(T_AmazonYoteiDate_Entity amazonyoteidate)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();

            dic.Add("@OperatorCD", amazonyoteidate.OperatorCD);
            dic.Add("@SEQ", amazonyoteidate.SEQ);
            dic.Add("@amazon_order_id", amazonyoteidate.amazon_order_id);
            dic.Add("@sku", amazonyoteidate.sku);
            return InsertUpdateDeleteData(dic, "T_AmazonYoteiDate_Update");
        }

        public DataTable T_AmazonYoteiDate_Select(T_AmazonYoteiDate_Entity amazonyoteidate)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@OperatorCD", amazonyoteidate.OperatorCD);
            dic.Add("@estimated_ship_date", amazonyoteidate.estimated_ship_date);
            return SelectData(dic, "T_AmazonYoteiDate_Select");
        }

    }
}
