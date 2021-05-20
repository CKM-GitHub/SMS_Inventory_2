using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
namespace SMS
{
   public class SplitDT
    {
       public List<DataTable> ListData(DataTable dt, int Rowss)
       {
           var co = dt.Rows.Count / Rowss;
           var remainer = dt.Rows.Count % Rowss;
           List<DataTable> dts = new List<DataTable>();

           for (int i = 0; i < co; i++)
           {
               var drw = dt.AsEnumerable().Skip(Rowss * i).Take(Rowss).CopyToDataTable(); ;
               dts.Add(drw);
           }

           if (remainer > 0)
           {
               var drw = dt.AsEnumerable().Skip(Rowss * co).Take(remainer).CopyToDataTable(); ;
               dts.Add(drw);
           }

           return dts;
       }
    }
}
