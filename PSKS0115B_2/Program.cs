using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Threading;
using SMS_BL;
using SMS_Entity;

namespace PSKS0115B_2//販売情報反映
{
    class Program
    {
        public static void Main(string[] args)
        {
            PSKS0115B_BL psks0115B_bl = new PSKS0115B_BL();
            DataTable dt = new DataTable();
            dt = psks0115B_bl.M_MultiPorpose_Select("109");

            while (dt.Rows[0]["Num1"].ToString().Equals("0"))
            {
                Console.WriteLine("Stop");
                Thread.Sleep(Convert.ToInt32(dt.Rows[0]["Num2"]) * 1000);
                dt = psks0115B_bl.M_MultiPorpose_Select("109");
            }

            Console.WriteLine("Start Import!");
            psks0115B_bl.PSKS0115B2_TZaiko_INSERT_UPDATE();
        }
    }
}
