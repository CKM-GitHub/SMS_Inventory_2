using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SMS_DL
{
  public  class PETC0310I_DL :Base_DL
    {
        public DataSet PETC310I_Select()
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@DateTime", DateTime.Now.ToString());
            return SelectDS(dic, "PETC310I_Select");
        }
    }
}
